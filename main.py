import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

os.environ['TK_SILENCE_DEPRECATION'] = '1'

class OrderCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Калькулятор заказов")
        self.root.geometry("350x250")
        self.root.resizable(False, False)
        
        self.center_window()
        
        self.data_file = ""
        self.limits_file = ""
        self.output_file = "результат.xlsx"
        
        self.data_selected = False
        self.limits_selected = False
        
        self.setup_ui()
        
    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
    def setup_ui(self):
        main_frame = tk.Frame(self.root, padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.data_status = tk.Label(main_frame, text="Файл data не выбран", 
                                   font=('Arial', 8), fg='red')
        self.data_status.pack(anchor='w', pady=(0, 5))
        
        data_btn = tk.Button(main_frame, 
                            text="Выбрать data.xlsx", 
                            command=self.browse_data_file,
                            font=('Arial', 10),
                            width=20)
        data_btn.pack(pady=3)
        
        self.limits_status = tk.Label(main_frame, text="Файл limits не выбран", 
                                     font=('Arial', 8), fg='red')
        self.limits_status.pack(anchor='w', pady=(5, 5))
        
        limits_btn = tk.Button(main_frame, 
                              text="Выбрать limits.xlsx", 
                              command=self.browse_limits_file,
                              font=('Arial', 10),
                              width=20)
        limits_btn.pack(pady=3)
        
        output_btn = tk.Button(main_frame, 
                              text="Имя файла результата", 
                              command=self.set_output_filename,
                              font=('Arial', 10),
                              width=20,
                              bg='lightgray')
        output_btn.pack(pady=8)
        
        self.filename_label = tk.Label(main_frame, 
                                      text=f"Результат: {self.output_file}",
                                      font=('Arial', 8),
                                      fg='blue')
        self.filename_label.pack(pady=2)
        
        self.process_btn = tk.Button(main_frame, 
                                    text="Рассчитать заказы", 
                                    command=self.process_data,
                                    font=('Arial', 10, 'bold'),
                                    bg='#4CAF50',
                                    fg='white',
                                    width=20,
                                    state='disabled')
        self.process_btn.pack(pady=10)
    
    def update_process_button_state(self):
        if self.data_selected and self.limits_selected:
            self.process_btn.config(state='normal', bg='#4CAF50')
        else:
            self.process_btn.config(state='disabled', bg='gray')
    
    def set_output_filename(self):
        filename = filedialog.asksaveasfilename(
            title="Сохранить результат как...",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=self.output_file
        )
        if filename:
            self.output_file = filename
            display_name = os.path.basename(filename)
            self.filename_label.config(text=f"Результат: {display_name}")
    
    def browse_data_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл с данными",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.data_file = filename
            self.data_selected = True
            display_name = os.path.basename(filename)
            self.data_status.config(text=f"✓ {display_name}", fg='green')
            self.update_process_button_state()
    
    def browse_limits_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл с лимитами",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.limits_file = filename
            self.limits_selected = True
            display_name = os.path.basename(filename)
            self.limits_status.config(text=f"✓ {display_name}", fg='green')
            self.update_process_button_state()
    
    def find_best_match(self, product_name, limits_dict):
        """
        Находит наиболее подходящий лимит для товара.
        Выбирает вариант с наибольшим количеством совпадающих символов.
        """
        best_match = None
        best_match_length = 0
        
        for limit_item in limits_dict.keys():
            # Если название товара содержит полное название лимита
            if limit_item in product_name:
                # Выбираем самый длинный совпадающий лимит
                if len(limit_item) > best_match_length:
                    best_match = limit_item
                    best_match_length = len(limit_item)
        
        return best_match
    
    def process_data(self):
        try:
            self.process_btn.config(state='disabled', bg='gray', text="Обработка...")
            self.root.update()
            
            # Загрузка данных
            df = pd.read_excel(self.data_file)
            limits_df = pd.read_excel(self.limits_file, sheet_name="Limits")
            blacklist_df = pd.read_excel(self.limits_file, sheet_name="BlackList")
            
            # Обрабатываем пустые значения в столбце "Остаток"
            # Заменяем NaN, пустые строки и другие нечисловые значения на 0
            df['Остаток'] = pd.to_numeric(df['Остаток'], errors='coerce').fillna(0)
            
            # Создаем словари
            limits = dict(zip(limits_df['Товар'].astype(str), limits_df['Лимиты']))
            blacklist = set(blacklist_df['Товар'].astype(str))
            
            # Фильтруем черный список
            df['Товар'] = df['Товар'].astype(str)
            df = df[~df['Товар'].isin(blacklist)]
            
            # Функция для расчета заказа с точным сопоставлением
            def calculate_order(row):
                product_name = row['Товар']
                best_match = self.find_best_match(product_name, limits)
                
                if best_match:
                    limit_value = limits[best_match]
                    return max(0, limit_value - row['Остаток'])
                return 0
            
            # Применяем расчет
            df['Заказ'] = df.apply(calculate_order, axis=1)
            
            # Добавляем столбец с использованным лимитом
            df['Использованный_лимит'] = df['Товар'].apply(
                lambda x: self.find_best_match(x, limits)
            )
            df['Лимиты'] = df['Использованный_лимит'].apply(
                lambda x: limits[x] if x else 0
            )
            
            # Удаляем временный столбец
            df = df.drop('Использованный_лимит', axis=1)
            
            # Сохраняем результат
            df.to_excel(self.output_file, index=False)
            
            messagebox.showinfo("Успех", 
                               f"Обработка завершена успешно!\n\n"
                               f"Файл сохранен как:\n{self.output_file}")
            
        except FileNotFoundError as e:
            messagebox.showerror("Ошибка файла", f"Файл не найден:\n{str(e)}")
        except KeyError as e:
            messagebox.showerror("Ошибка данных", f"Отсутствует необходимая колонка:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла непредвиденная ошибка:\n{str(e)}")
        
        finally:
            self.process_btn.config(state='normal', bg='#4CAF50', text="Рассчитать заказы")
            self.update_process_button_state()

def main():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    
    root = tk.Tk()
    app = OrderCalculatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()