import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re

os.environ['TK_SILENCE_DEPRECATION'] = '1'

class OrderCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Калькулятор заказов - Пакетная обработка")
        self.root.geometry("500x400")
        self.root.resizable(False, False)
        
        self.center_window()
        
        self.base_folder = ""
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
        
        title_label = tk.Label(main_frame, 
                              text="Пакетная обработка торговых точек", 
                              font=('Arial', 12, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # Выбор папки с автозаявками
        folder_frame = tk.Frame(main_frame)
        folder_frame.pack(fill=tk.X, pady=5)
        
        self.folder_status = tk.Label(folder_frame, text="Папка с автозаявками не выбрана", 
                                     font=('Arial', 8), fg='red', wraplength=450)
        self.folder_status.pack(anchor='w')
        
        folder_btn = tk.Button(folder_frame, 
                              text="Выбрать папку с автозаявками", 
                              command=self.browse_base_folder,
                              font=('Arial', 10),
                              width=25)
        folder_btn.pack(pady=5)
        
        # Информация о найденных торговых точках
        info_frame = tk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=10)
        
        self.info_label = tk.Label(info_frame, 
                                  text="Торговые точки не найдены",
                                  font=('Arial', 9),
                                  justify=tk.LEFT)
        self.info_label.pack(anchor='w')
        
        # Прогресс бар
        progress_frame = tk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=450, mode='determinate')
        self.progress.pack(pady=5)
        
        self.progress_label = tk.Label(progress_frame, text="", font=('Arial', 8))
        self.progress_label.pack()
        
        # Кнопка обработки
        self.process_btn = tk.Button(main_frame, 
                                    text="Обработать все торговые точки", 
                                    command=self.process_all_stores,
                                    font=('Arial', 10, 'bold'),
                                    width=25,
                                    state='disabled')
        self.process_btn.pack(pady=15)
        
        # Лог обработки
        log_frame = tk.Frame(main_frame)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        log_label = tk.Label(log_frame, text="Лог обработки:", font=('Arial', 9, 'bold'))
        log_label.pack(anchor='w')
        
        self.log_text = tk.Text(log_frame, height=8, width=60, font=('Arial', 8))
        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def browse_base_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку с автозаявками")
        if folder:
            self.base_folder = folder
            display_name = os.path.basename(folder)
            self.folder_status.config(text=f"✓ {folder}", fg='green')
            
            # Поиск торговых точек
            stores = self.find_stores(folder)
            if stores:
                self.process_btn.config(state='normal', bg='#4CAF50')
                self.info_label.config(text=f"Найдено торговых точек: {len(stores)}\nПапки: {', '.join(stores)}")
            else:
                self.process_btn.config(state='disabled', bg='gray')
                self.info_label.config(text="Торговые точки не найдены")
    
    def find_stores(self, base_folder):
        """Находит все папки с торговыми точками в базовой папке"""
        stores = []
        for item in os.listdir(base_folder):
            item_path = os.path.join(base_folder, item)
            if os.path.isdir(item_path):
                # Проверяем, есть ли в папке нужные файлы
                data_file = os.path.join(item_path, "Data.xlsx")
                limits_file = os.path.join(item_path, "limits.xlsx")
                
                if os.path.exists(data_file) and os.path.exists(limits_file):
                    stores.append(item)
        
        return stores
    
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
    
    def process_store(self, store_folder, store_name):
        """Обрабатывает одну торговую точку"""
        try:
            data_file = os.path.join(store_folder, "Data.xlsx")
            limits_file = os.path.join(store_folder, "limits.xlsx")
            output_file = os.path.join(self.base_folder, f"{store_name}.xlsx")
            
            # Загрузка данных
            df = pd.read_excel(data_file)
            limits_df = pd.read_excel(limits_file, sheet_name="Limits")
            blacklist_df = pd.read_excel(limits_file, sheet_name="BlackList")
            
            # Обрабатываем пустые значения в столбце "Остаток"
            df['Остаток'] = pd.to_numeric(df['Остаток'], errors='coerce').fillna(0)
            
            # Создаем словари
            limits = dict(zip(limits_df['Товар'].astype(str), limits_df['Лимиты']))
            blacklist = set(blacklist_df['Товар'].astype(str))
            
            # Фильтруем черный список - улучшенная версия
            df['Товар'] = df['Товар'].astype(str)

            # Создаем маску с использованием границ слов
            def is_blacklisted(product):
                for black_item in blacklist:
                    # Ищем black_item как отдельное слово
                    pattern = r'\b' + re.escape(black_item) + r'\b'
                    if re.search(pattern, product, re.IGNORECASE):
                        return True
                return False

            # Оставляем только строки, которые НЕ подпадают под черный список
            mask = df['Товар'].apply(is_blacklisted)
            df = df[~mask]
            
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
            df.to_excel(output_file, index=False)
            
            return True, f"✓ {store_name} - Успешно"
            
        except Exception as e:
            return False, f"✗ {store_name} - Ошибка: {str(e)}"
    
    def log_message(self, message):
        """Добавляет сообщение в лог"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def process_all_stores(self):
        """Обрабатывает все торговые точки"""
        if not self.base_folder:
            messagebox.showerror("Ошибка", "Сначала выберите папку с автозаявками")
            return
        
        stores = self.find_stores(self.base_folder)
        if not stores:
            messagebox.showerror("Ошибка", "Торговые точки не найдены")
            return
        
        # Настройка интерфейса для обработки
        self.process_btn.config(state='disabled', bg='gray', text="Обработка...")
        self.log_text.delete(1.0, tk.END)
        self.progress['value'] = 0
        self.progress['maximum'] = len(stores)
        
        # Обработка каждой торговой точки
        success_count = 0
        for i, store in enumerate(stores):
            self.progress_label.config(text=f"Обработка: {store} ({i+1}/{len(stores)})")
            self.progress['value'] = i
            
            store_folder = os.path.join(self.base_folder, store)
            success, message = self.process_store(store_folder, store)
            
            self.log_message(message)
            if success:
                success_count += 1
            
            self.root.update()
        
        # Завершение
        self.progress['value'] = len(stores)
        self.progress_label.config(text="Обработка завершена")
        self.process_btn.config(state='normal', bg='#4CAF50', text="Обработать все торговые точки")
        
        messagebox.showinfo("Завершено", 
                          f"Обработка завершена!\n\n"
                          f"Успешно: {success_count} из {len(stores)}")

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