import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from excel_handler import get_column_max_lengths
from settings import Settings  # Импортируем Settings

class SettingsDialog:
    def __init__(self, parent):
        self.parent = parent
        self.settings = Settings()
        self.dialog = None
        self.save_path_var = None
        self.columns_var = None
        
    def show(self):
        """Показать диалоговое окно настроек"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Settings")
        self.dialog.geometry("600x400")
        self.dialog.resizable(False, False)
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Центрирование окна
        self.dialog.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() - self.dialog.winfo_width()) // 2
        y = self.parent.winfo_y() + (self.parent.winfo_height() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        self.create_widgets()
        self.load_current_settings()
        
    def create_widgets(self):
        """Создать виджеты диалогового окна"""
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Настраиваем веса строк для правильного распределения пространства
        main_frame.rowconfigure(3, weight=1)  # Даем строке с чекбоксами возможность растягиваться
        main_frame.rowconfigure(4, weight=0)  # Кнопки фиксированы внизу
        
        # Путь сохранения
        ttk.Label(main_frame, text="Save Path:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        path_frame = ttk.Frame(main_frame)
        path_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        path_frame.columnconfigure(0, weight=1)
        
        self.save_path_var = tk.StringVar()
        path_entry = ttk.Entry(path_frame, textvariable=self.save_path_var)
        path_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        browse_btn = ttk.Button(path_frame, text="Browse", command=self.browse_save_path)
        browse_btn.grid(row=0, column=1)
        
        # Столбцы для сохранения
        ttk.Label(main_frame, text="Columns to Keep:").grid(
            row=2, column=0, sticky=tk.W, pady=(0, 5))

        # Фрейм для чекбоксов с прокруткой - ОГРАНИЧИВАЕМ ВЫСОТУ
        checkbox_frame = ttk.Frame(main_frame)
        checkbox_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Добавляем скроллбар с фиксированной высотой
        canvas = tk.Canvas(checkbox_frame, height=120)  # Фиксированная высота
        scrollbar = ttk.Scrollbar(checkbox_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Создаем чекбоксы для колонок 1-40
        self.checkbox_vars = {}
        Setts_Columns = (
            ("ISIN"),
            ("Ticker & Exchange"), 
            ("Ccy"),
            ("Cpn (%)"),
            ("Name"),
            ("Sector"),
            ("Industry"), 
            ("Maturity (1. call date)"),
            ("Price"),
            ("Perf YTD %"),
            ("Mk-Cap mia"),
            ("YTM MID"),
            ("Share classes"),
            ("ER/MF"),
            ("Rating Moody"),
            ("Rating S&P"),
            ("Rating Fitch"),
            ("Size mio"),
            ("Z- Spread"),
            ("ASW spread"),
            ("Min piece"),
            ("Min incr"),
            ("Mkt of Issue"),
            ("Notes"),
            ("Added on")
        )
        for i in range(1, 26):
            var = tk.BooleanVar()
            self.checkbox_vars[i] = var
            checkbox = ttk.Checkbutton(scrollable_frame, text=str(Setts_Columns[i-1]), variable=var)
            checkbox.grid(row=(i-1)//4, column=(i-1)%4, sticky=tk.W, padx=5, pady=2)

        # Упаковываем с заполнением
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Кнопки - ПЕРЕМЕЩАЕМ В ОТДЕЛЬНУЮ СТРОКУ
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, sticky=tk.E, pady=(10, 0))
        
        cancel_btn = ttk.Button(button_frame, text="Cancel", command=self.dialog.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        save_btn = ttk.Button(button_frame, text="Save", command=self.save_settings)
        save_btn.pack(side=tk.RIGHT)
        
        # Настройка весов для растягивания
        main_frame.columnconfigure(0, weight=1)
        
    def load_current_settings(self):
        """Загрузить текущие настройки в поля"""
        self.save_path_var.set(self.settings.get_save_path())
        
        # Устанавливаем чекбоксы согласно сохраненным настройкам
        columns_to_keep = self.settings.get_column_to_keep()
        for col_num, var in self.checkbox_vars.items():
            var.set(col_num in columns_to_keep)
        
    def browse_save_path(self):
        """Выбрать путь для сохранения"""
        path = filedialog.askdirectory(
            title="Select Save Directory",
            initialdir=self.save_path_var.get()
        )
        if path:
            self.save_path_var.set(path)
            
    def save_settings(self):
        """Сохранить настройки"""
        try:
            save_path = self.save_path_var.get().strip()
            
            # Валидация пути
            if not save_path:
                messagebox.showerror("Error", "Save path cannot be empty")
                return
                
            # Получаем выбранные колонки из чекбоксов
            columns_to_keep = []
            for col_num, var in self.checkbox_vars.items():
                if var.get():
                    columns_to_keep.append(col_num)
            
            # Валидация столбцов
            if not columns_to_keep:
                messagebox.showerror("Error", "At least one column must be selected")
                return
            
            # Сохранение настроек - ИСПРАВЛЕНО ИМЯ МЕТОДА
            self.settings.save_settings(save_path, columns_to_keep)
            messagebox.showinfo("Success", "Settings saved successfully")
            self.dialog.destroy()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")

# Остальной код класса ExcelAppGUI остается без изменений...
class ExcelAppGUI:
    def __init__(self, root, on_file_load, on_open_settings):
        self.root = root
        self.root.title("Light App")
        self.root.geometry("800x600")

        self.on_file_load = on_file_load
        self.on_open_settings = on_open_settings
        
        self.current_file_path = None
        self.sheet_data = {}
        self.current_sheet = None
        self.selected_rows = {}

        self.create_widgets()
        self.sheet_listbox.insert(tk.END, "No file loaded")
        self.sheet_listbox.config(state=tk.DISABLED)

    def get_selected_rows(self):
        """Возвращает словарь с выделенными строками"""
        return self.selected_rows

    def get_sheet_data(self):
        """Возвращает словарь с данными листов"""
        return self.sheet_data

    def get_current_file_path(self):
        """Возвращает путь к текущему файлу"""
        return self.current_file_path

    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Toolbar
        toolbar = ttk.Frame(main_frame)
        toolbar.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Open file button
        open_btn = ttk.Button(toolbar, text="Open File", command=self.open_file_dialog_handler)
        open_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # Settings button
        settings_btn = ttk.Button(toolbar, text="Settings", command=self.open_settings)
        settings_btn.pack(side=tk.LEFT)
        
        # Sheet selection frame
        sheet_frame = ttk.Frame(main_frame)
        sheet_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10), padx=(0, 10))
        sheet_frame.columnconfigure(0, weight=1)
        sheet_frame.rowconfigure(1, weight=1)

        ttk.Label(sheet_frame, text="Sheets:").grid(row=0, column=0, sticky=tk.W)

        # Frame для listbox и scrollbar
        listbox_frame = ttk.Frame(sheet_frame)
        listbox_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        listbox_frame.columnconfigure(0, weight=1)
        listbox_frame.rowconfigure(0, weight=1)

        self.sheet_listbox = tk.Listbox(listbox_frame, width=20)
        self.sheet_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Scrollbar для списка листов
        listbox_scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.sheet_listbox.yview)
        listbox_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        self.sheet_listbox.configure(yscrollcommand=listbox_scrollbar.set)
        self.sheet_listbox.bind('<<ListboxSelect>>', self.on_sheet_select)
        
        # Table frame
        table_frame = ttk.Frame(main_frame)
        table_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        
        # Create treeview with scrollbars - ТОЛЬКО ОДИН БИНДИНГ
        self.tree = ttk.Treeview(table_frame, show='headings', selectmode='extended')
        self.tree.bind('<Button-1>', self.on_tree_click)  # ← ТОЛЬКО ЭТОТ

        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Select a file to begin")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=1, column=0, sticky=(tk.W, tk.E))
    
    def open_file_dialog_handler(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.load_excel_file(file_path)
    
    def load_excel_file(self, file_path):
        try:
            self.status_var.set("Loading file...")
            self.root.update()
            
            sheet_data = self.on_file_load(file_path)
            
            if not sheet_data:
                messagebox.showerror("Error", "Failed to load Excel file or file is empty")
                return
            
            self.current_file_path = file_path
            self.sheet_data = sheet_data
            
            self.sheet_listbox.config(state=tk.NORMAL)
            self.sheet_listbox.delete(0, tk.END)
            for sheet_name in sheet_data.keys():
                self.sheet_listbox.insert(tk.END, sheet_name)
            
            if self.sheet_listbox.size() > 0:
                self.sheet_listbox.selection_set(0)
                self.display_sheet(list(sheet_data.keys())[0])
            
            self.status_var.set(f"Loaded: {file_path}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")
            self.status_var.set("Error loading file")

    def display_sheet(self, sheet_name):
        self.current_sheet = sheet_name
        df = self.sheet_data[sheet_name]
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Используем 2-ю и 3-ю строки как заголовки, начиная с 3-го столбца
        if len(df) >= 3:
            # Объединяем заголовки из 2-й и 3-й строк, начиная с 3-го столбца
            columns = []
            for i in range(2, len(df.columns)):  # Начинаем с 3-го столбца (индекс 2)
                header1 = str(df.iloc[1, i]) if len(df) > 1 and pd.notna(df.iloc[1, i]) else ""
                header2 = str(df.iloc[2, i]) if len(df) > 2 and pd.notna(df.iloc[2, i]) else ""
                
                # Объединяем заголовки через пробел, если оба не пустые
                if header1 and header2:
                    column_name = f"{header1} {header2}"
                else:
                    column_name = header1 or header2 or f"Column {i+1}"
                
                columns.append(column_name)
        else:
            # Если недостаточно строк, используем обычные заголовки, начиная с 3-го столбца
            columns = list(df.columns[2:])
        
        self.tree["columns"] = columns
        
        max_lengths = {}
        for col_idx, col_name in enumerate(columns):
            # Вычисляем максимальную длину для каждого столбца, начиная с 5-й строки
            if len(df) >= 5:
                max_len = df.iloc[3:, col_idx + 2].astype(str).apply(len).max()  # +2 для смещения к 3-му столбцу
                max_lengths[col_name] = max(max_len, len(col_name)) if not pd.isna(max_len) else len(col_name)
            else:
                max_lengths[col_name] = len(col_name)

        for col in columns:
            self.tree.heading(col, text=col)
            col_width = max(max_lengths.get(col, 10) * 8, len(col) * 8)
            self.tree.column(col, width=col_width, minwidth=col_width)
        
        # Отображаем данные начиная с 5-й строки (индекс 4) и с 3-го столбца (индекс 2)
        start_row = 3 if len(df) >= 4 else 0
        for i in range(start_row, len(df)):
            values = [str(df.iloc[i, col_idx + 2]) if pd.notna(df.iloc[i, col_idx + 2]) else "" for col_idx in range(len(columns))]
            item_id = f"{sheet_name}_{i + 2}"
            self.tree.insert("", "end", values=values, iid=item_id)

        if sheet_name in self.selected_rows and self.selected_rows[sheet_name]:
            for row_idx in self.selected_rows[sheet_name]:
                item_id = f"{sheet_name}_{row_idx}"
                if self.tree.exists(item_id):
                    self.tree.selection_add(item_id)

    def on_tree_click(self, event):
        """Обрабатывает клик мыши для toggle выделения"""
        if not self.current_sheet:
            return
        
        item = self.tree.identify_row(event.y)
        if not item:
            return
        
        current_selection = set(self.tree.selection())
        
        if item in current_selection:
            current_selection.remove(item)
        else:
            current_selection.add(item)
        
        self.tree.selection_set(list(current_selection))
        self.update_selected_rows()
        
        return "break"

    def update_selected_rows(self):
        if not self.current_sheet:
            return
        selected_items = self.tree.selection()
        selected_indices = []
        for item in selected_items:
            try:
                row_idx = int(item.split('_')[-1])  # Без +2, реальный индекс
                selected_indices.append(row_idx)
            except (ValueError, IndexError):
                continue
        self.selected_rows[self.current_sheet] = selected_indices

    def on_sheet_select(self, event):
        if not self.sheet_listbox.curselection():
            return
        
        self.update_selected_rows()
        selected_index = self.sheet_listbox.curselection()[0]
        sheet_name = self.sheet_listbox.get(selected_index)
        self.display_sheet(sheet_name)

    def open_settings(self):
        """Открыть диалоговое окно настроек"""
        settings_dialog = SettingsDialog(self.root)
        settings_dialog.show()

    def run(self):
        self.root.mainloop()

