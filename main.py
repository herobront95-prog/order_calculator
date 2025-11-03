import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re

class BlackListFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Фильтр по черному списку")
        self.root.geometry("600x500")
        self.root.resizable(False, False)
        
        self.center_window()
        self.setup_ui()
        
        self.data_file = ""
        self.blacklist_file = ""
        
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
                              text="Фильтр файлов по черному списку", 
                              font=('Arial', 12, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # Выбор основного файла
        file_frame = tk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(file_frame, text="Основной файл:", font=('Arial', 9, 'bold')).pack(anchor='w')
        
        self.file_status = tk.Label(file_frame, text="Файл не выбран", 
                                   font=('Arial', 8), fg='red', wraplength=450)
        self.file_status.pack(anchor='w')
        
        file_btn = tk.Button(file_frame, 
                           text="Выбрать основной файл", 
                           command=self.browse_data_file,
                           font=('Arial', 10),
                           width=20)
        file_btn.pack(pady=5)
        
        # Выбор файла черного списка
        blacklist_frame = tk.Frame(main_frame)
        blacklist_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(blacklist_frame, text="Файл черного списка:", font=('Arial', 9, 'bold')).pack(anchor='w')
        
        self.blacklist_status = tk.Label(blacklist_frame, text="Файл не выбран", 
                                       font=('Arial', 8), fg='red', wraplength=450)
        self.blacklist_status.pack(anchor='w')
        
        blacklist_btn = tk.Button(blacklist_frame, 
                                text="Выбрать черный список", 
                                command=self.browse_blacklist_file,
                                font=('Arial', 10),
                                width=20)
        blacklist_btn.pack(pady=5)
        
        # Информация о процессе
        info_frame = tk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=10)
        
        self.info_label = tk.Label(info_frame, text="", font=('Arial', 9), wraplength=450)
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
                                   text="Обработать файл", 
                                   command=self.process_file,
                                   font=('Arial', 10, 'bold'),
                                   width=20,
                                   state='disabled',
                                   bg='lightgray')
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
    
    def browse_data_file(self):
        """Выбор основного файла для обработки"""
        file = filedialog.askopenfilename(
            title="Выберите файл для обработки",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file:
            self.data_file = file
            display_name = os.path.basename(file)
            self.file_status.config(text=f"✓ {display_name}", fg='green')
            self.update_process_button()
    
    def browse_blacklist_file(self):
        """Выбор файла с черным списком"""
        file = filedialog.askopenfilename(
            title="Выберите файл черного списка",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file:
            self.blacklist_file = file
            display_name = os.path.basename(file)
            self.blacklist_status.config(text=f"✓ {display_name}", fg='green')
            self.update_process_button()
    
    def update_process_button(self):
        """Активирует кнопку обработки если оба файла выбраны"""
        if self.data_file and self.blacklist_file:
            self.process_btn.config(state='normal', bg='#4CAF50')
            self.info_label.config(text="Оба файла выбраны. Можно начинать обработку.")
        else:
            self.process_btn.config(state='disabled', bg='lightgray')
    
    def load_blacklist(self):
        """Загружает черный список из файла"""
        try:
            # Пробуем прочитать файл
            blacklist_df = pd.read_excel(self.blacklist_file)
            
            # Берем первый столбец и преобразуем в список
            first_column_name = blacklist_df.columns[0]
            blacklist_items = blacklist_df[first_column_name].dropna().astype(str).tolist()
            
            self.log_message(f"Загружено {len(blacklist_items)} позиций из черного списка")
            return blacklist_items
            
        except Exception as e:
            self.log_message(f"Ошибка загрузки черного списка: {str(e)}")
            return []
    
    def find_best_match(self, product_name, blacklist):
        """
        Находит наиболее подходящее соответствие в черном списке.
        Все слова из черного списка должны присутствовать в названии товара.
        """
        best_match = None
        best_match_score = 0
        
        for black_item in blacklist:
            # Разбиваем элемент черного списка на слова
            words = [word.strip() for word in black_item.split() if word.strip()]
            
            if not words:
                continue
                
            # Создаем паттерн для проверки что ВСЕ слова присутствуют
            patterns = []
            for word in words:
                escaped_word = re.escape(word)
                patterns.append(f"(?=.*{escaped_word})")
            
            full_pattern = "".join(patterns) + ".*"
            
            # Проверяем соответствие
            if re.search(full_pattern, product_name, re.IGNORECASE):
                # Оцениваем по количеству слов и длине совпадения
                score = len(words) * 1000 + len(black_item)
                if score > best_match_score:
                    best_match = black_item
                    best_match_score = score
        
        return best_match
    
    def log_message(self, message):
        """Добавляет сообщение в лог"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def process_file(self):
        """Основная функция обработки"""
        if not self.data_file or not self.blacklist_file:
            messagebox.showerror("Ошибка", "Выберите оба файла")
            return
        
        # Настройка интерфейса для обработки
        self.process_btn.config(state='disabled', bg='gray', text="Обработка...")
        self.log_text.delete(1.0, tk.END)
        self.progress['value'] = 0
        
        try:
            # Загрузка черного списка
            self.log_message("=== ЗАГРУЗКА ЧЕРНОГО СПИСКА ===")
            blacklist = self.load_blacklist()
            
            if not blacklist:
                messagebox.showerror("Ошибка", "Не удалось загрузить черный список")
                self.process_btn.config(state='normal', bg='#4CAF50', text="Обработать файл")
                return
            
            # Загрузка основного файла
            self.log_message("=== ЗАГРУЗКА ОСНОВНОГО ФАЙЛА ===")
            df = pd.read_excel(self.data_file)
            initial_count = len(df)
            self.log_message(f"Загружено {initial_count} строк")
            
            # Преобразуем первый столбец в строки
            first_column_name = df.columns[0]
            df[first_column_name] = df[first_column_name].astype(str)
            
            # Поиск совпадений и удаление
            self.log_message("=== ПОИСК СОВПАДЕНИЙ ===")
            self.progress['maximum'] = len(df)
            
            rows_to_remove = []
            removed_items = []
            
            for i, row in df.iterrows():
                product_name = row[first_column_name]
                match = self.find_best_match(product_name, blacklist)
                
                if match:
                    rows_to_remove.append(i)
                    removed_items.append((product_name, match))
                
                if i % 10 == 0:  # Обновляем прогресс каждые 10 строк
                    self.progress['value'] = i
                    self.progress_label.config(text=f"Обработано: {i}/{len(df)} строк")
                    self.root.update()
            
            # Удаляем найденные строки
            if rows_to_remove:
                df_cleaned = df.drop(rows_to_remove)
                removed_count = len(rows_to_remove)
            else:
                df_cleaned = df
                removed_count = 0
            
            # Сохранение результата
            self.log_message("=== СОХРАНЕНИЕ РЕЗУЛЬТАТА ===")
            output_file = self.get_output_filename()
            df_cleaned.to_excel(output_file, index=False)
            
            # Логируем результаты
            self.log_message(f"\n=== РЕЗУЛЬТАТЫ ===")
            self.log_message(f"Исходное количество строк: {initial_count}")
            self.log_message(f"Удалено строк: {removed_count}")
            self.log_message(f"Осталось строк: {len(df_cleaned)}")
            self.log_message(f"Файл сохранен как: {os.path.basename(output_file)}")
            
            if removed_items:
                self.log_message(f"\nУдаленные позиции:")
                for product, match in removed_items[:10]:  # Показываем первые 10
                    self.log_message(f"  '{product}' → совпадение с '{match}'")
                if len(removed_items) > 10:
                    self.log_message(f"  ... и еще {len(removed_items) - 10} позиций")
            
            # Завершение
            self.progress['value'] = len(df)
            self.progress_label.config(text="Обработка завершена")
            self.process_btn.config(state='normal', bg='#4CAF50', text="Обработать файл")
            
            messagebox.showinfo("Завершено", 
                              f"Обработка завершена!\n\n"
                              f"Удалено: {removed_count} из {initial_count} строк\n"
                              f"Сохранено в: {os.path.basename(output_file)}")
            
        except Exception as e:
            self.log_message(f"Ошибка при обработке: {str(e)}")
            messagebox.showerror("Ошибка", f"Произошла ошибка при обработке: {str(e)}")
            self.process_btn.config(state='normal', bg='#4CAF50', text="Обработать файл")
    
    def get_output_filename(self):
        """Генерирует имя для выходного файла"""
        base_name = os.path.splitext(self.data_file)[0]
        counter = 1
        output_file = f"{base_name}_filtered.xlsx"
        
        # Если файл уже существует, добавляем номер
        while os.path.exists(output_file):
            output_file = f"{base_name}_filtered_{counter}.xlsx"
            counter += 1
        
        return output_file

def main():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    
    root = tk.Tk()
    app = BlackListFilterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()