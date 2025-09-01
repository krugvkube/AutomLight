from gui import ExcelAppGUI
import tkinter as tk
from tkinter import ttk, messagebox
from excel_handler import load_excel_data, excel_processing
from settings import Settings  # Добавлен импорт Settings

def main():
    root = tk.Tk()
    app = ExcelAppGUI(root, on_file_load=load_excel_data, on_open_settings=None)
    settings = Settings()  # Создаем экземпляр настроек
    
    # Создаем кнопку для обработки данных
    toolbar = app.root.nametowidget('.!frame.!frame')  # Получаем доступ к toolbar
    process_btn = ttk.Button(toolbar, text="Process", command=lambda: process_excel_data(app, settings))
    process_btn.pack(side=tk.LEFT, padx=(5, 0))
    
    app.run()

def process_excel_data(app, settings):
    """Обработка выделенных данных Excel"""
    if not app.current_file_path:
        messagebox.showerror("Error", "No file loaded")
        return
        
    if not app.selected_rows:
        messagebox.showerror("Error", "No rows selected")
        return
        
    try:
        # Передаем данные в функцию обработки
        excel_processing(
            file_path=app.current_file_path,
            sheet_data=app.sheet_data,
            selected_rows=app.selected_rows,
            save_path=settings.get_save_path(),
            columns_to_keep=settings.get_column_to_keep()
        )

        messagebox.showinfo("Success", "Processing completed successfully")
    except Exception as e:
        messagebox.showerror("Error", f"Processing failed: {str(e)}")

if __name__ == "__main__":
    main()
