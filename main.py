import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
import sys

os.environ['TK_SILENCE_DEPRECATION'] = '1'

class OrderCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Калькулятор заказов")
        self.root.geometry("300x180")
        self.root.resizable(False, False)
        
        self.center_window()
        
        self.data_file = ""
        self.limits_file = ""
        self.output_file = "результат.xlsx"
        
        self.setup_ui()
        
    def center_window(self):
        self.root.update_idletasks()
        width = 300
        height = 180
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
    def setup_ui(self):
        main_frame = tk.Frame(self.root, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Кнопка для выбора файла с данными
        data_btn = tk.Button(main_frame, 
                            text="Выбрать data.xlsx", 
                            command=self.browse_data_file,
                            font=('Arial', 10),
                            width=20)
        data_btn.pack(pady=3)
        
        # Кнопка для выбора файла с лимитами
        limits_btn = tk.Button(main_frame, 
                              text="Выбрать limits.xlsx", 
                              command=self.browse_limits_file,
                              font=('Arial', 10),
                              width=20)
        limits_btn.pack(pady=3)
        
        # Кнопка для установки имени выходного файла
        output_btn = tk.Button(main_frame, 
                              text="Имя файла результата", 
                              command=self.set_output_filename,
                              font=('Arial', 10),
                              width=20,
                              bg='lightgray')
        output_btn.pack(pady=3)
        
        # Метка для отображения текущего имени файла
        self.filename_label = tk.Label(main_frame, 
                                      text=f"Будет сохранен как: {self.output_file}",
                                      font=('Arial', 8),
                                      fg='blue')
        self.filename_label.pack(pady=2)
        
        # Кнопка для расчета заказов
        self.process_btn = tk.Button(main_frame, 
                                    text="Рассчитать заказы", 
                                    command=self.process_data,
                                    font=('Arial', 10, 'bold'),
                                    bg='#4CAF50',
                                    fg='white',
                                    width=20)
        self.process_btn.pack(pady=5)
    
    def set_output_filename(self):
        filename = filedialog.asksaveasfilename(
        title="Сохранить результат как...",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile=self.output_file
        )
        if filename:
            self.output_file = filename
            self.filename_label.config(text=f"Будет сохранен как: {self.output_file}")
    
    def browse_data_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл с данными",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.data_file = filename
            self.process_btn.config(text="Файл data выбран!")
            self.root.after(2000, lambda: self.process_btn.config(text="Рассчитать заказы"))
    
    def browse_limits_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл с лимитами",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.limits_file = filename
            self.process_btn.config(text="Файл limits выбран!")
            self.root.after(2000, lambda: self.process_btn.config(text="Рассчитать заказы"))
    
    def process_data(self):
        # Проверяем, что файлы выбраны
        if not self.data_file:
            messagebox.showerror("Ошибка", "Пожалуйста, сначала выберите файл data.xlsx")
            return
        
        if not self.limits_file:
            messagebox.showerror("Ошибка", "Пожалуйста, сначала выберите файл limits.xlsx")
            return
        
        try:
            self.process_btn.config(state='disabled', bg='gray', text="Обработка...")
            
            # Загрузка и обработка данных
            df = pd.read_excel(self.data_file)
            limits_df = pd.read_excel(self.limits_file, sheet_name="Limits")
            blacklist_df = pd.read_excel(self.limits_file, sheet_name="BlackList")
            
            limits = dict(zip(limits_df['Товар'], limits_df['Лимиты']))
            blacklist = blacklist_df['Товар'].astype(str).tolist()
            
            df['Товар'] = df['Товар'].astype(str)
            df = df[~df['Товар'].apply(lambda x: any(name in x for name in blacklist))]
            
            def calculate(row):
                for item, rate in limits.items():
                    if item in row['Товар']:
                        return max(0, rate - row['Остаток'])
                return 0
            
            df['Заказ'] = df.apply(calculate, axis=1)
            df['Лимиты'] = df.apply(lambda row: 
                                   next((rate for item, rate in limits.items() 
                                         if item in row['Товар']), 0), axis=1)
            
            # Сохранение результата
            df.to_excel(self.output_file, index=False)
            
            messagebox.showinfo("Успех", 
                               f"Обработка завершена!\n\n"
                               f"Результат сохранен в файл:\n{self.output_file}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")
        
        finally:
            self.process_btn.config(state='normal', bg='#4CAF50', text="Рассчитать заказы")

def main():
    root = tk.Tk()
    app = OrderCalculatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()