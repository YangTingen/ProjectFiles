import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import time
import random
import sqlite3

class CNCControlInterface:
    def __init__(self, root):
        self.root = root
        self.root.title("CNC 四軸機械手臂控制")
        # 設置全螢幕
        self.root.attributes('-fullscreen', True)
        self.root.configure(bg="#2F2F2F")

        # 當前坐標 (X, Y, Z, C)
        self.coords = {"X": 0.0, "Y": 0.0, "Z": 0.0, "C": 0.0}
        self.is_running = False  # 自動模式下程式執行狀態
        self.is_paused = False  # 暫停狀態
        self.current_file = None  # 儲存當前讀取的檔案路徑
        self.operation_mode = "手動"  # 操作模式：手動 或 自動
        self.execution_mode = "連續"  # 執行模式：連續 或 單節（自動模式下使用）
        self.current_line = 0  # 當前執行行數
        self.total_lines = 0  # 總行數

        # 移動距離選項（下拉式選單）
        self.move_distances = ["0.01", "0.1", "0.5", "1.0", "5.0", "10.0"]
        self.move_distance = tk.StringVar(value="1.0")  # 預設移動距離為 1.0

        # 元件狀態 (OUTPUT 和 INPUT 將從資料庫動態載入)
        self.output_components = {}  # OUTPUT: 根據 io 表格動態生成
        self.input_components = {}   # INPUT: 根據 io 表格動態生成
        self.output_buttons = {}  # 儲存 OUTPUT 按鈕的引用
        self.input_labels = {}    # 儲存 INPUT 標籤的引用
        self.axis_buttons = []    # 儲存軸控制按鈕的引用
        self.auto_buttons = []    # 儲存自動模式按鈕的引用（啟動、暫停、停止）

        # 資料庫初始化
        self.db_name = "machine_data.db"
        self.init_database()

        # 資料表編輯狀態
        self.edited_rows = set()  # 儲存被編輯但未儲存的行（IID）
        self.original_data = {}   # 儲存原始資料，用於恢復

        # 定義樣式
        self.configure_styles()

        # 檢測元件初始狀態
        self.check_initial_state()

        # 創建主框架
        self.create_widgets()

    def init_database(self):
        # 初始化資料庫，如果表格不存在則建立
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                # 建立 point 表格，如果不存在
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS point (
                        name TEXT PRIMARY KEY,
                        x REAL,
                        y REAL,
                        z REAL,
                        c REAL
                    )
                ''')
                # 檢查 io 表格是否存在並確認其結構
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS io (
                        name TEXT PRIMARY KEY,
                        io TEXT,
                        number INTEGER
                    )
                ''')
                # 檢查 io 表格是否有資料，如果沒有則插入預設資料（可選）
                cursor.execute("SELECT COUNT(*) FROM io")
                count = cursor.fetchone()[0]
                if count == 0:
                    # 插入預設資料（範例）
                    default_data = [
                        ("感測器1", "input", 6),
                        ("感測器2", "input", 7),
                        ("感測器3", "input", 8),
                        ("氣缸1", "output", 1),
                        ("氣缸2", "output", 2),
                        ("氣缸3", "output", 3),
                    ]
                    cursor.executemany("INSERT INTO io (name, io, number) VALUES (?, ?, ?)", default_data)
                conn.commit()
                print("資料庫初始化完成")
        except sqlite3.Error as e:
            print(f"資料庫初始化失敗: {e}")
            messagebox.showerror("錯誤", f"無法初始化資料庫: {e}")

    def configure_styles(self):
        # 使用 ttk.Style 定義按鈕樣式
        style = ttk.Style()
        style.theme_use('default')

        # 設置整體背景和字體
        style.configure("TFrame", background="#2F2F2F")
        style.configure("TLabel", background="#2F2F2F", foreground="white", font=("Helvetica", 10))
        style.configure("TLabelFrame", background="#2F2F2F", foreground="white", font=("Helvetica", 12))
        style.configure("TLabelFrame.Label", background="#2F2F2F", foreground="white")

        # 通用按鈕樣式（矩形、細邊框）
        style.configure("TButton", font=("Helvetica", 10, "bold"), borderwidth=1, relief="solid", padding=5, background="#D3D3D3", foreground="black")
        style.map("TButton",
                  background=[("active", "#4A90E2"), ("disabled", "#808080")],
                  foreground=[("disabled", "#A0A0A0")])

        # 軸控制按鈕
        style.configure("Axis.TButton", background="#D3D3D3", foreground="black")
        style.map("Axis.TButton",
                  background=[("active", "#4A90E2"), ("disabled", "#808080")])

        # OUTPUT 按鈕（使用顏色區分開關狀態）
        style.configure("OutputOn.TButton", background="#00FF00", foreground="black", font=("Helvetica", 9, "bold"), borderwidth=1, relief="solid", padding=5)
        style.map("OutputOn.TButton",
                  background=[("active", "#00CC00"), ("disabled", "#00FF00")],
                  foreground=[("disabled", "#A0A0A0")])
        style.configure("OutputOff.TButton", background="#FF0000", foreground="black", font=("Helvetica", 9, "bold"), borderwidth=1, relief="solid", padding=5)
        style.map("OutputOff.TButton",
                  background=[("active", "#CC0000"), ("disabled", "#FF0000")],
                  foreground=[("disabled", "#A0A0A0")])

        # INPUT 狀態標籤
        style.configure("StateOn.TLabel", background="#00FF00", foreground="black", font=("Helvetica", 9, "bold"), borderwidth=1, relief="solid", padding=5)
        style.configure("StateOff.TLabel", background="#FF0000", foreground="black", font=("Helvetica", 9, "bold"), borderwidth=1, relief="solid", padding=5)

        # 自動模式按鈕（啟動、暫停、停止）
        style.configure("Start.TButton", background="#00FF00", foreground="black")
        style.map("Start.TButton",
                  background=[("active", "#00CC00"), ("disabled", "#808080")])
        style.configure("Pause.TButton", background="#FFA500", foreground="black")
        style.map("Pause.TButton",
                  background=[("active", "#CC8400"), ("disabled", "#808080")])
        style.configure("Stop.TButton", background="#FF0000", foreground="black")
        style.map("Stop.TButton",
                  background=[("active", "#CC0000"), ("disabled", "#808080")])

        # 模式切換按鈕
        style.configure("Mode.TButton", background="#D3D3D3", foreground="black")
        style.map("Mode.TButton",
                  background=[("active", "#4A90E2"), ("disabled", "#808080")])

        # 關閉程式按鈕
        style.configure("Close.TButton", background="#D3D3D3", foreground="black")
        style.map("Close.TButton",
                  background=[("active", "#4A90E2"), ("disabled", "#808080")])

        # 程式碼編輯按鈕
        style.configure("File.TButton", background="#D3D3D3", foreground="black")
        style.map("File.TButton",
                  background=[("active", "#4A90E2"), ("disabled", "#808080")])

        # 儲存編輯按鈕（動態樣式）
        style.configure("SaveNormal.TButton", background="#D3D3D3", foreground="black", font=("Helvetica", 10, "bold"), borderwidth=1, relief="solid", padding=5)
        style.map("SaveNormal.TButton",
                  background=[("active", "#4A90E2"), ("disabled", "#808080")],
                  foreground=[("disabled", "#A0A0A0")])
        style.configure("SaveEdited.TButton", background="#D3D3D3", foreground="#FF0000", font=("Helvetica", 10, "bold"), borderwidth=1, relief="solid", padding=5)
        style.map("SaveEdited.TButton",
                  background=[("active", "#4A90E2"), ("disabled", "#808080")],
                  foreground=[("disabled", "#A0A0A0")])

        # 下拉式選單樣式
        style.configure("TCombobox", fieldbackground="#263238", background="#D3D3D3", foreground="black")
        style.map("TCombobox", fieldbackground=[("readonly", "#263238")], selectbackground=[("readonly", "#263238")])

        # 資料表樣式
        style.configure("Treeview", background="#263238", foreground="white", fieldbackground="#263238")
        style.configure("Treeview.Heading", background="#D3D3D3", foreground="black", font=("Helvetica", 10, "bold"))

    def check_initial_state(self):
        # 從 io 表格讀取資料，根據 io 欄位區分 input 和 output
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT name, io FROM io")
                rows = cursor.fetchall()

                # 清空現有的 input 和 output 狀態
                self.input_components.clear()
                self.output_components.clear()

                # 根據 io 欄位分類
                for name, io_type in rows:
                    if io_type.lower() == "input":
                        self.input_components[name] = random.choice([True, False])
                    elif io_type.lower() == "output":
                        self.output_components[name] = False  # 預設關閉

                print("OUTPUT 元件初始狀態:", self.output_components)
                print("INPUT 元件初始狀態:", self.input_components)
        except sqlite3.Error as e:
            print(f"無法從 io 表格讀取資料: {e}")
            messagebox.showerror("錯誤", f"無法從 io 表格讀取資料: {e}")

    def create_widgets(self):
        # 主框架，使用 PanedWindow 來分隔左右區域
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill="both", expand=True, padx=5, pady=5)

        # 左側框架：資訊顯示區域
        left_frame = ttk.Frame(main_paned)
        main_paned.add(left_frame, weight=2)

        # 右側框架：控制按鈕區域
        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=1)

        # 左側：標題
        title_label = tk.Label(left_frame, text="CNC 四軸機械手臂控制", font=("Helvetica", 14, "bold"), bg="#2F2F2F", fg="white")
        title_label.pack(pady=5)

        # 左側：坐標顯示框架（表格形式）
        coord_frame = ttk.LabelFrame(left_frame, text="坐標")
        coord_frame.pack(pady=5, fill="x")
        coord_inner_frame = tk.Frame(coord_frame, bg="#2F2F2F")
        coord_inner_frame.pack(pady=5, padx=5)

        self.coord_labels = {}
        for axis in ["X", "Y", "Z", "C"]:
            frame = tk.Frame(coord_inner_frame, bg="#2F2F2F")
            frame.pack(fill="x", pady=2)
            tk.Label(frame, text=f"{axis}:", width=5, font=("Helvetica", 10), bg="#2F2F2F", fg="white").pack(side=tk.LEFT)
            unit = "°" if axis == "C" else "mm"
            display_text = f"0.000 {unit}"
            self.coord_labels[axis] = tk.Label(frame, text=display_text, font=("Helvetica", 10), bg="#2F2F2F", fg="#00FF00", width=15)
            self.coord_labels[axis].pack(side=tk.LEFT)

        # 左側：程式碼顯示框架（高度為 10 行）
        code_frame = ttk.LabelFrame(left_frame, text="程式碼")
        code_frame.pack(pady=5, fill="both", expand=True)
        self.code_text = tk.Text(code_frame, height=10, width=50, font=("Helvetica", 10), bg="#263238", fg="white", insertbackground="white")
        self.code_text.pack(pady=5, padx=5, fill="both", expand=True)

        # 程式碼操作按鈕
        code_button_frame = tk.Frame(code_frame, bg="#2F2F2F")
        code_button_frame.pack(pady=5)
        ttk.Button(code_button_frame, text="讀取檔案", style="File.TButton", command=self.load_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(code_button_frame, text="儲存檔案", style="File.TButton", command=self.save_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(code_button_frame, text="另存新檔", style="File.TButton", command=self.save_file_as).pack(side=tk.LEFT, padx=5)

        # 左側：自動控制（1x3 格子）
        auto_frame = ttk.LabelFrame(left_frame, text="自動控制")
        auto_frame.pack(pady=5, fill="x")
        auto_grid = tk.Frame(auto_frame, bg="#2F2F2F")
        auto_grid.pack(pady=5)

        start_button = ttk.Button(auto_grid, text="啟動", width=10, style="Start.TButton", command=self.start_machine)
        start_button.grid(row=0, column=0, padx=5, pady=5)
        pause_button = ttk.Button(auto_grid, text="暫停", width=10, style="Pause.TButton", command=self.pause_machine)
        pause_button.grid(row=0, column=1, padx=5, pady=5)
        stop_button = ttk.Button(auto_grid, text="停止", width=10, style="Stop.TButton", command=self.stop_machine)
        stop_button.grid(row=0, column=2, padx=5, pady=5)
        self.auto_buttons = [start_button, pause_button, stop_button]

        # 左側：模式切換（1x3 格子，寬度加寬為 16）
        mode_frame = ttk.LabelFrame(left_frame, text="模式切換")
        mode_frame.pack(pady=5, fill="x")
        mode_grid = tk.Frame(mode_frame, bg="#2F2F2F")
        mode_grid.pack(pady=5)

        self.mode_button = ttk.Button(mode_grid, text=f"操作模式: {self.operation_mode}", width=16, style="Mode.TButton", command=self.toggle_operation_mode)
        self.mode_button.grid(row=0, column=0, padx=5, pady=5)
        self.exec_mode_button = ttk.Button(mode_grid, text=f"執行模式: {self.execution_mode}", width=16, style="Mode.TButton", command=self.toggle_execution_mode)
        self.exec_mode_button.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(mode_grid, text="關閉程式", width=10, style="Close.TButton", command=self.close_program).grid(row=0, column=2, padx=5, pady=5)

        # 左側：狀態和進度顯示框架
        status_frame = ttk.LabelFrame(left_frame, text="狀態資訊")
        status_frame.pack(pady=5, fill="x")
        self.status_label = tk.Label(status_frame, text="狀態: 已停止", font=("Helvetica", 10), bg="#2F2F2F", fg="#00FF00")
        self.status_label.pack(pady=5)
        self.progress_label = tk.Label(status_frame, text="進度: 0/0", font=("Helvetica", 10), bg="#2F2F2F", fg="#00FF00")
        self.progress_label.pack(pady=5)

        # 右側：控制按鈕區域
        # 手動控制：方向鍵 (2x4 格子) + 移動距離下拉選單
        manual_frame = ttk.LabelFrame(right_frame, text="手動控制")
        manual_frame.pack(pady=5, fill="x")
        manual_inner_frame = tk.Frame(manual_frame, bg="#2F2F2F")
        manual_inner_frame.pack(pady=5)

        # 移動距離下拉選單
        distance_frame = tk.Frame(manual_inner_frame, bg="#2F2F2F")
        distance_frame.pack(fill="x", pady=5)
        tk.Label(distance_frame, text="移動距離:", width=10, font=("Helvetica", 10), bg="#2F2F2F", fg="white").pack(side=tk.LEFT)
        self.distance_combobox = ttk.Combobox(distance_frame, textvariable=self.move_distance, values=self.move_distances, width=8)
        self.distance_combobox.pack(side=tk.LEFT, padx=5)

        # 方向鍵
        manual_grid = tk.Frame(manual_inner_frame, bg="#2F2F2F")
        manual_grid.pack(pady=5)

        axis_controls = [
            ("X+", "X", 1), ("X-", "X", -1),
            ("Y+", "Y", 1), ("Y-", "Y", -1),
            ("Z+", "Z", 1), ("Z-", "Z", -1),
            ("C+", "C", 1), ("C-", "C", -1),
        ]
        for idx, (text, axis, direction) in enumerate(axis_controls):
            button = ttk.Button(manual_grid, text=text, width=8, style="Axis.TButton", command=lambda a=axis, d=direction: self.move_axis(a, d))
            button.grid(row=idx//2, column=idx%2, padx=5, pady=5)
            self.axis_buttons.append(button)

        # OUTPUT 控制：動態生成按鈕，根據 io 表格
        output_frame = ttk.LabelFrame(right_frame, text="OUTPUT 控制")
        output_frame.pack(pady=5, fill="x")
        output_grid = tk.Frame(output_frame, bg="#2F2F2F")
        output_grid.pack(pady=5)

        # 根據 output_components 動態生成按鈕
        output_names = list(self.output_components.keys())
        for idx, comp_name in enumerate(output_names):
            state = self.output_components[comp_name]
            button_style = "OutputOn.TButton" if state else "OutputOff.TButton"
            button = ttk.Button(output_grid, text=comp_name, width=10, style=button_style, command=lambda c=comp_name: self.toggle_output(c))
            button.grid(row=idx//3, column=idx%3, padx=5, pady=5)
            self.output_buttons[comp_name] = button

        # INPUT 狀態：動態生成標籤，根據 io 表格
        input_frame = ttk.LabelFrame(right_frame, text="INPUT 狀態")
        input_frame.pack(pady=5, fill="x")
        input_grid = tk.Frame(input_frame, bg="#2F2F2F")
        input_grid.pack(pady=5)

        # 根據 input_components 動態生成標籤
        input_names = list(self.input_components.keys())
        for idx, comp_name in enumerate(input_names):
            state = self.input_components[comp_name]
            label_style = "StateOn.TLabel" if state else "StateOff.TLabel"
            label = ttk.Label(input_grid, text=comp_name, width=10, style=label_style)
            label.grid(row=idx//3, column=idx%3, padx=5, pady=5)
            self.input_labels[comp_name] = label

        # 右側：資料表（顯示 point 表格）
        data_frame = ttk.LabelFrame(right_frame, text="資料表")
        data_frame.pack(pady=5, fill="both", expand=True)

        # 創建一個框架來放置 Treeview 和滾動條
        table_frame = tk.Frame(data_frame, bg="#2F2F2F")
        table_frame.pack(fill="both", expand=True)

        # 使用 Treeview 顯示資料表
        columns = ("name", "x", "y", "z", "c")
        self.data_table = ttk.Treeview(table_frame, columns=columns, show="headings", height=5)
        self.data_table.heading("name", text="名稱")
        self.data_table.heading("x", text="X (mm)")
        self.data_table.heading("y", text="Y (mm)")
        self.data_table.heading("z", text="Z (mm)")
        self.data_table.heading("c", text="C (°)")
        self.data_table.column("name", width=100, anchor="center")
        self.data_table.column("x", width=80, anchor="center")
        self.data_table.column("y", width=80, anchor="center")
        self.data_table.column("z", width=80, anchor="center")
        self.data_table.column("c", width=80, anchor="center")

        # 創建滾動條
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.data_table.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        self.data_table.configure(yscrollcommand=scrollbar.set)
        self.data_table.pack(fill="both", expand=True)

        # 綁定雙擊事件以編輯資料表
        self.data_table.bind("<Double-1>", self.on_double_click)

        # 資料表操作按鈕
        data_button_frame = tk.Frame(data_frame, bg="#2F2F2F")
        data_button_frame.pack(pady=5)
        ttk.Button(data_button_frame, text="更新", style="File.TButton", command=self.refresh_data_table).pack(side=tk.LEFT, padx=5)
        ttk.Button(data_button_frame, text="新增", style="File.TButton", command=self.add_new_data).pack(side=tk.LEFT, padx=5)
        self.save_button = ttk.Button(data_button_frame, text="儲存編輯", style="SaveNormal.TButton", command=self.save_edited_data)
        self.save_button.pack(side=tk.LEFT, padx=5)
        ttk.Button(data_button_frame, text="刪除", style="File.TButton", command=self.delete_data).pack(side=tk.LEFT, padx=5)
        move_button = ttk.Button(data_button_frame, text="移動到位置", style="File.TButton")
        move_button.pack(side=tk.LEFT, padx=5)
        # 綁定按鈕點擊事件以檢查 Ctrl 鍵
        move_button.bind("<Button-1>", self.move_to_selected_position)

        # 初次載入資料
        self.refresh_data_table()

        # 初始模式：手動模式
        self.update_button_states()

    def on_double_click(self, event):
        # 雙擊編輯資料表單元格
        item = self.data_table.selection()
        if not item:
            return
        item = item[0]
        column = self.data_table.identify_column(event.x)
        col_index = int(column.replace("#", "")) - 1
        if col_index >= 5:  # 防止越界
            return

        # 獲取當前值
        current_value = self.data_table.item(item, "values")[col_index]

        # 創建輸入框
        entry = ttk.Entry(self.data_table)
        entry.insert(0, current_value)
        entry.place(x=event.x, y=event.y, anchor="nw")

        def save_edit(event):
            new_value = entry.get()
            values = list(self.data_table.item(item, "values"))
            # 驗證數值欄位
            if col_index > 0:  # x, y, z, c 必須是數字
                try:
                    new_value = float(new_value)
                except ValueError:
                    messagebox.showwarning("警告", "請輸入有效的數字！")
                    entry.destroy()
                    return
            values[col_index] = new_value
            self.data_table.item(item, values=values)
            self.edited_rows.add(item)
            # 儲存原始資料
            if item not in self.original_data:
                self.original_data[item] = list(self.data_table.item(item, "values"))
            entry.destroy()
            self.update_control_states()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.focus_set()

    def add_new_data(self):
        # 新增一筆資料，以當前坐標為預設值
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                # 計算現有資料筆數，生成名稱 "Point_<序號>"
                cursor.execute("SELECT COUNT(*) FROM point")
                count = cursor.fetchone()[0]
                name = f"Point_{count + 1}"

                # 插入資料
                cursor.execute('''
                    INSERT INTO point (name, x, y, z, c) VALUES (?, ?, ?, ?, ?)
                ''', (name, self.coords["X"], self.coords["Y"], self.coords["Z"], self.coords["C"]))
                conn.commit()
                print(f"已新增資料: {name}, X={self.coords['X']}, Y={self.coords['Y']}, Z={self.coords['Z']}, C={self.coords['C']}")
                messagebox.showinfo("提示", f"已新增資料: {name}")
                # 刷新資料表
                self.refresh_data_table()
        except sqlite3.Error as e:
            print(f"無法新增資料: {e}")
            messagebox.showerror("錯誤", f"無法新增資料: {e}")

    def save_edited_data(self):
        # 將編輯後的資料儲存到資料庫
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                valid_items = []
                for item in list(self.edited_rows):
                    # 檢查 item 是否仍存在於 Treeview 中
                    if item in self.data_table.get_children():
                        values = self.data_table.item(item, "values")
                        name = values[0]
                        x, y, z, c = map(float, values[1:5])
                        cursor.execute('''
                            UPDATE point SET x = ?, y = ?, z = ?, c = ? WHERE name = ?
                        ''', (x, y, z, c, name))
                        valid_items.append(item)
                conn.commit()
                print("編輯資料已儲存到資料庫")
                messagebox.showinfo("提示", "編輯資料已儲存")
                # 清除編輯狀態
                self.edited_rows.clear()
                self.original_data.clear()
                self.refresh_data_table()
        except sqlite3.Error as e:
            print(f"無法儲存編輯資料: {e}")
            messagebox.showerror("錯誤", f"無法儲存編輯資料: {e}")

    def delete_data(self):
        # 刪除選中的資料
        selected_item = self.data_table.selection()
        if not selected_item:
            messagebox.showwarning("警告", "請選擇要刪除的資料！")
            return

        item = selected_item[0]
        values = self.data_table.item(item, "values")
        name = values[0]

        if messagebox.askyesno("確認", f"確定要刪除資料 '{name}' 嗎？"):
            try:
                with sqlite3.connect(self.db_name) as conn:
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM point WHERE name = ?", (name,))
                    conn.commit()
                    print(f"已刪除資料: {name}")
                    messagebox.showinfo("提示", f"已刪除資料: {name}")
                    # 刷新資料表
                    self.refresh_data_table()
            except sqlite3.Error as e:
                print(f"無法刪除資料: {e}")
                messagebox.showerror("錯誤", f"無法刪除資料: {e}")

    def refresh_data_table(self):
        # 從資料庫讀取資料並更新資料表
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT name, x, y, z, c FROM point")
                rows = cursor.fetchall()

                # 清空現有資料
                for item in self.data_table.get_children():
                    self.data_table.delete(item)

                # 更新資料表
                valid_edited_rows = set()
                for row in rows:
                    name = row[0]
                    values = list(row)
                    item = self.data_table.insert("", tk.END, values=values)
                    if item in self.edited_rows:
                        valid_edited_rows.add(item)

                # 更新 edited_rows，移除不存在的 item
                self.edited_rows = valid_edited_rows
                print("資料表已更新")
                self.update_control_states()
        except sqlite3.Error as e:
            print(f"無法讀取資料庫: {e}")
            messagebox.showerror("錯誤", f"無法讀取資料庫: {e}")

    def move_to_position(self, item):
        # 獲取資料表中的位置
        values = self.data_table.item(item, "values")
        name = values[0]
        try:
            target_coords = {
                "X": float(values[1]),
                "Y": float(values[2]),
                "Z": float(values[3]),
                "C": float(values[4]),
            }
        except (ValueError, IndexError) as e:
            messagebox.showwarning("警告", f"資料格式錯誤，無法移動到位置 '{name}'！錯誤：{e}")
            print(f"移動失敗，資料格式錯誤: {e}")
            return False

        # 計算移動距離（用於模擬移動時間）
        total_distance = 0.0
        for axis in ["X", "Y", "Z", "C"]:
            distance = abs(target_coords[axis] - self.coords[axis])
            total_distance += distance

        # 模擬移動時間（每 1 mm 或 1° 花費 0.1 秒）
        move_time = total_distance * 0.1
        if move_time < 0.5:  # 最小移動時間為 0.5 秒
            move_time = 0.5

        # 更新狀態顯示為「移動中」
        self.status_label.config(text=f"狀態: 移動中 ({name})")
        self.root.update()  # 立即更新介面
        print(f"開始移動到位置 '{name}': X={target_coords['X']}, Y={target_coords['Y']}, Z={target_coords['Z']}, C={target_coords['C']}")

        # 模擬移動
        time.sleep(move_time)

        # 更新坐標
        for axis, value in target_coords.items():
            self.coords[axis] = value
            unit = "°" if axis == "C" else "mm"
            display_text = f"{self.coords[axis]:.3f} {unit}"
            self.coord_labels[axis].config(text=display_text)

        # 更新狀態顯示為「已停止」
        self.status_label.config(text="狀態: 已停止")
        print(f"移動完成: X={self.coords['X']}, Y={self.coords['Y']}, Z={self.coords['Z']}, C={self.coords['C']}")
        messagebox.showinfo("提示", f"已移動到位置: {name}")
        return True

    def move_to_selected_position(self, event):
        # 檢查是否按住 Ctrl 鍵
        if not (event.state & 0x4):  # 0x4 表示 Ctrl 鍵
            messagebox.showinfo("提示", "請按住 Ctrl 鍵以移動到指定位置！")
            return

        # 檢查是否有選中的資料表行
        selected_item = self.data_table.selection()
        if selected_item:
            # 如果有選中的行，直接移動到該位置
            success = self.move_to_position(selected_item[0])
            if not success:
                print("移動失敗，可能是資料格式錯誤")
            return

        # 如果沒有選中的行，顯示下拉選單選擇位置
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM point")
                position_names = [row[0] for row in cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"無法讀取資料庫: {e}")
            messagebox.showerror("錯誤", f"無法讀取資料庫: {e}")
            return

        if not position_names:
            messagebox.showwarning("警告", "目前沒有可用的位置！")
            return

        # 創建一個頂層視窗來顯示下拉選單
        dialog = tk.Toplevel(self.root)
        dialog.title("選擇移動位置")
        dialog.configure(bg="#2F2F2F")
        dialog.transient(self.root)
        dialog.grab_set()

        tk.Label(dialog, text="選擇要移動到的位置：", font=("Helvetica", 10), bg="#2F2F2F", fg="white").pack(pady=5)
        position_var = tk.StringVar(value=position_names[0])
        position_combobox = ttk.Combobox(dialog, textvariable=position_var, values=position_names, state="readonly")
        position_combobox.pack(pady=5)

        def confirm_move():
            selected_name = position_var.get()
            # 查找對應的 item
            for item in self.data_table.get_children():
                values = self.data_table.item(item, "values")
                if values[0] == selected_name:
                    success = self.move_to_position(item)
                    if not success:
                        print(f"移動到位置 '{selected_name}' 失敗")
                    break
            dialog.destroy()

        ttk.Button(dialog, text="確認", command=confirm_move).pack(pady=5)
        ttk.Button(dialog, text="取消", command=dialog.destroy).pack(pady=5)

        # 居中顯示對話框
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")

    def update_control_states(self):
        # 根據編輯狀態啟用或禁用控制功能，並更新「儲存編輯」按鈕的樣式
        if self.edited_rows:
            # 有未儲存的編輯，禁用手動和自動控制，按鈕文字變為紅色
            for button in self.axis_buttons:
                button.config(state=tk.DISABLED)
            for button in self.output_buttons.values():
                button.config(state=tk.DISABLED)
            for button in self.auto_buttons:
                button.config(state=tk.DISABLED)
            self.exec_mode_button.config(state=tk.DISABLED)
            self.distance_combobox.config(state="disabled")
            self.mode_button.config(state=tk.DISABLED)
            self.save_button.config(style="SaveEdited.TButton")  # 文字變為紅色
        else:
            # 無未儲存的編輯，根據操作模式更新按鈕狀態，按鈕文字恢復黑色
            self.update_button_states()
            self.save_button.config(style="SaveNormal.TButton")  # 文字恢復黑色

    def toggle_operation_mode(self):
        # 限制模式切換：必須在程式停止時才能切換
        if self.is_running:
            messagebox.showwarning("警告", "程式正在運行或暫停中，請先停止程式再切換模式！")
            return
        # 切換操作模式：手動 或 自動
        self.operation_mode = "自動" if self.operation_mode == "手動" else "手動"
        self.mode_button.config(text=f"操作模式: {self.operation_mode}")
        self.update_button_states()
        print(f"操作模式切換為: {self.operation_mode}")

    def toggle_execution_mode(self):
        # 切換執行模式：連續 或 單節（僅在自動模式下生效）
        self.execution_mode = "單節" if self.execution_mode == "連續" else "連續"
        self.exec_mode_button.config(text=f"執行模式: {self.execution_mode}")
        print(f"執行模式切換為: {self.execution_mode}")

    def update_button_states(self):
        # 根據操作模式啟用或禁用按鈕
        if self.operation_mode == "手動":
            # 手動模式：啟用軸控制和 OUTPUT 控制按鈕，禁用自動模式按鈕
            for button in self.axis_buttons:
                button.config(state=tk.NORMAL)
            for button in self.output_buttons.values():
                button.config(state=tk.NORMAL)
            for button in self.auto_buttons:
                button.config(state=tk.DISABLED)
            self.exec_mode_button.config(state=tk.DISABLED)
            self.distance_combobox.config(state="normal")
        else:
            # 自動模式：啟用自動模式按鈕，禁用軸控制和 OUTPUT 控制按鈕
            for button in self.axis_buttons:
                button.config(state=tk.DISABLED)
            for button in self.output_buttons.values():
                button.config(state=tk.DISABLED)
            for button in self.auto_buttons:
                button.config(state=tk.NORMAL)
            self.exec_mode_button.config(state=tk.NORMAL)
            self.distance_combobox.config(state="disabled")
        self.mode_button.config(state=tk.NORMAL)

    def close_program(self):
        # 關閉程式
        self.root.destroy()

    def load_file(self):
        # 選擇檔案並讀取
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("CNC files", "*.cnc")])
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    # 清空文字欄並顯示檔案內容
                    self.code_text.delete(1.0, tk.END)
                    self.code_text.insert(tk.END, file.read())
                self.current_file = file_path
                print(f"已讀取檔案: {file_path}")
            except Exception as e:
                messagebox.showerror("錯誤", f"無法讀取檔案: {e}")

    def save_file(self):
        # 儲存到原檔案
        if not self.current_file:
            messagebox.showwarning("警告", "請先讀取一個檔案或使用另存新檔!")
            return
        try:
            with open(self.current_file, 'w', encoding='utf-8') as file:
                file.write(self.code_text.get(1.0, tk.END).strip())
            print(f"已儲存到檔案: {self.current_file}")
            messagebox.showinfo("提示", f"檔案已儲存到: {self.current_file}")
        except Exception as e:
            messagebox.showerror("錯誤", f"無法儲存檔案: {e}")

    def save_file_as(self):
        # 另存為新檔案
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", 
                                                 filetypes=[("Text files", "*.txt"), ("CNC files", "*.cnc")])
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(self.code_text.get(1.0, tk.END).strip())
                self.current_file = file_path
                print(f"已另存為檔案: {file_path}")
                messagebox.showinfo("提示", f"檔案已儲存到: {file_path}")
            except Exception as e:
                messagebox.showerror("錯誤", f"無法儲存檔案: {e}")

    def toggle_output(self, component):
        # 切換 OUTPUT 元件狀態（僅改變顏色，不改變文字）
        if self.operation_mode != "手動":
            return
        self.output_components[component] = not self.output_components[component]
        style = "OutputOn.TButton" if self.output_components[component] else "OutputOff.TButton"
        self.output_buttons[component].config(style=style)
        print(f"{component} 現在狀態: {'ON' if self.output_components[component] else 'OFF'}")

    def move_axis(self, axis, direction):
        # 僅在手動模式下允許移動軸
        if self.operation_mode != "手動":
            return
        # 獲取移動距離
        try:
            distance = float(self.move_distance.get())
            if distance <= 0:
                raise ValueError("移動距離必須大於 0")
        except ValueError:
            messagebox.showwarning("警告", "請輸入有效的移動距離（正數）")
            self.move_distance.set("1.0")  # 恢復預設值
            return
        # 模擬移動，按選擇的距離移動
        self.coords[axis] += direction * distance
        unit = "°" if axis == "C" else "mm"
        display_text = f"{self.coords[axis]:.3f} {unit}"
        self.coord_labels[axis].config(text=display_text)
        print(f"移動 {axis} 軸到 {display_text}")

    def start_machine(self):
        if self.operation_mode != "自動":
            return
        if self.is_running:
            if self.is_paused:
                # 從暫停狀態繼續執行
                self.is_paused = False
                self.status_label.config(text="狀態: 運行中")
                print("機械手臂繼續執行")
                self.execute_next_line(self.current_line)
                return
            messagebox.showwarning("警告", "機械手臂已在運行!")
            return
        self.is_running = True
        self.is_paused = False
        self.current_line = 0  # 重置行數
        self.status_label.config(text="狀態: 運行中")
        messagebox.showinfo("狀態", "機械手臂已啟動")
        print("機械手臂啟動")

        # 一行一行執行文字欄中的程式碼
        code_lines = self.code_text.get(1.0, tk.END).strip().splitlines()
        if not code_lines or not code_lines[0]:
            messagebox.showwarning("警告", "程式碼欄位為空，無法執行!")
            self.stop_machine()
            return

        self.code_lines = code_lines
        self.total_lines = len(self.code_lines)
        self.update_progress()  # 更新進度顯示
        self.execute_next_line(0)

    def pause_machine(self):
        if self.operation_mode != "自動":
            return
        if not self.is_running:
            messagebox.showwarning("警告", "機械手臂未在運行!")
            return
        if self.is_paused:
            messagebox.showwarning("警告", "機械手臂已在暫停狀態!")
            return
        self.is_paused = True
        self.status_label.config(text="狀態: 暫停中")
        print("機械手臂已暫停")

    def execute_next_line(self, index):
        if not self.is_running or self.is_paused or index >= len(self.code_lines):
            return
        # 高亮當前行
        self.code_text.tag_remove("highlight", "1.0", tk.END)
        self.code_text.tag_add("highlight", f"{index+1}.0", f"{index+1}.end")
        self.code_text.tag_config("highlight", background="#FFFF00", foreground="black")

        line = self.code_lines[index].strip()
        if line:
            print(f"執行指令: {line}")
        self.current_line = index + 1
        self.update_progress()  # 更新進度顯示

        if self.execution_mode == "連續":
            # 連續執行模式：自動執行下一行
            self.root.after(1000, self.execute_next_line, self.current_line)  # 每行延遲 1 秒
        else:
            # 單節執行模式：執行一行後暫停
            self.is_paused = True
            self.status_label.config(text="狀態: 暫停中 (單節執行)")
            print("單節執行完成，等待繼續")

        if self.current_line >= len(self.code_lines):
            self.stop_machine()

    def update_progress(self):
        # 更新進度顯示
        self.progress_label.config(text=f"進度: {self.current_line}/{self.total_lines}")

    def stop_machine(self):
        if self.operation_mode != "自動":
            return
        self.is_running = False
        self.is_paused = False
        self.current_line = 0
        self.total_lines = 0
        self.status_label.config(text="狀態: 已停止")
        self.update_progress()  # 重置進度顯示
        self.code_text.tag_remove("highlight", "1.0", tk.END)  # 移除高亮
        messagebox.showinfo("狀態", "機械手臂已停止")
        print("機械手臂停止")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = CNCControlInterface(root)
        root.mainloop()
    except tk.TclError as e:
        print(f"無法啟動圖形介面: {e}")
        print("請確保環境支援圖形顯示 (例如設置 $DISPLAY 變數或在本地電腦運行)。")