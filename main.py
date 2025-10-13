import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re

os.environ['TK_SILENCE_DEPRECATION'] = '1'

class CheckboxListFrame(tk.Frame):
    """Фрейм со списком чекбоксов и прокруткой"""
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.checkboxes = {}
        self.checkvars = {}
        
        # Создаем фрейм для канвы и скроллбара
        self.canvas_frame = tk.Frame(self)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        # Канва и скроллбар
        self.canvas = tk.Canvas(self.canvas_frame, height=150)
        self.scrollbar = ttk.Scrollbar(self.canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Фрейм для кнопок управления
        self.control_frame = tk.Frame(self)
        self.control_frame.pack(fill=tk.X, pady=5)
        
        self.select_all_btn = tk.Button(self.control_frame, text="Выбрать все", command=self.select_all)
        self.select_all_btn.pack(side=tk.LEFT, padx=5)
        
        self.deselect_all_btn = tk.Button(self.control_frame, text="Снять все", command=self.deselect_all)
        self.deselect_all_btn.pack(side=tk.LEFT, padx=5)
        
        self.selected_count_label = tk.Label(self.control_frame, text="Выбрано: 0")
        self.selected_count_label.pack(side=tk.RIGHT, padx=5)
    
    def add_checkbox(self, name, text):
        """Добавляет чекбокс в список"""
        var = tk.BooleanVar(value=True)
        self.checkvars[name] = var
        
        cb = ttk.Checkbutton(self.scrollable_frame, text=text, variable=var, 
                           command=self.update_count)
        cb.pack(anchor="w", padx=5, pady=2)
        self.checkboxes[name] = cb
        
        self.update_count()
    
    def select_all(self):
        """Выбирает все чекбоксы"""
        for var in self.checkvars.values():
            var.set(True)
        self.update_count()
    
    def deselect_all(self):
        """Снимает все чекбоксы"""
        for var in self.checkvars.values():
            var.set(False)
        self.update_count()
    
    def update_count(self):
        """Обновляет счетчик выбранных"""
        count = sum(1 for var in self.checkvars.values() if var.get())
        self.selected_count_label.config(text=f"Выбрано: {count}")
    
    def get_selected(self):
        """Возвращает список выбранных имен"""
        return [name for name, var in self.checkvars.items() if var.get()]
    
    def clear(self):
        """Очищает список"""
        for checkbox in self.checkboxes.values():
            checkbox.destroy()
        self.checkboxes.clear()
        self.checkvars.clear()
        self.update_count()

class OrderCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Калькулятор заказов - Пакетная обработка")
        self.root.geometry("600x700")
        self.root.resizable(False, False)
        
        self.center_window()
        
        self.base_folder = ""
        self.stores = []
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
                              text="Калькулятор заказов - Пакетная обработка", 
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
        
        # Кнопка диагностики и добавления лимитов
        control_buttons_frame = tk.Frame(folder_frame)
        control_buttons_frame.pack(pady=5)
        
        diagnose_btn = tk.Button(control_buttons_frame,
                                text="Диагностика папки",
                                command=self.diagnose_folder,
                                font=('Arial', 9),
                                width=15,
                                bg='lightblue')
        diagnose_btn.pack(side=tk.LEFT, padx=5)
        
        self.add_limit_btn = tk.Button(control_buttons_frame,
                                text="Добавить лимит",
                                command=self.open_add_limit_window,
                                font=('Arial', 9),
                                width=15,
                                bg='lightgreen',
                                state='disabled')
        self.add_limit_btn.pack(side=tk.LEFT, padx=5)
        
        # Список торговых точек с чекбоксами
        list_label = tk.Label(main_frame, text="Выберите торговые точки для обработки:", 
                             font=('Arial', 9, 'bold'))
        list_label.pack(anchor='w', pady=(10, 5))
        
        self.checkbox_frame = CheckboxListFrame(main_frame, relief=tk.SUNKEN, bd=1)
        self.checkbox_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Чекбокс "Сократить заявку"
        options_frame = tk.Frame(main_frame)
        options_frame.pack(fill=tk.X, pady=10)
        
        self.reduce_order_var = tk.BooleanVar()
        self.reduce_order_cb = tk.Checkbutton(
            options_frame, 
            text="Сократить заявку", 
            variable=self.reduce_order_var,
            font=('Arial', 9)
        )
        self.reduce_order_cb.pack(anchor='w')
        
        # Информация о сокращении
        self.reduce_info_label = tk.Label(
            options_frame, 
            text="При включении: удаляются позиции с заказом ≤ 1/3 от лимита и позиции с лимитом=2 и заказом=1",
            font=('Arial', 8),
            fg='gray',
            wraplength=450,
            justify=tk.LEFT
        )
        self.reduce_info_label.pack(anchor='w', pady=(2, 0))
        
        # Прогресс бар
        progress_frame = tk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=450, mode='determinate')
        self.progress.pack(pady=5)
        
        self.progress_label = tk.Label(progress_frame, text="", font=('Arial', 8))
        self.progress_label.pack()
        
        # Кнопка обработки
        self.process_btn = tk.Button(main_frame, 
                                    text="Обработать выбранные точки", 
                                    command=self.process_selected_stores,
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
            self.stores = self.find_stores(folder)
            self.update_checkbox_list()
            
            if self.stores:
                self.process_btn.config(state='normal', bg='#4CAF50')
                self.add_limit_btn.config(state='normal')
            else:
                self.process_btn.config(state='disabled', bg='gray')
                self.add_limit_btn.config(state='disabled')
    
    def update_checkbox_list(self):
        """Обновляет список чекбоксов"""
        self.checkbox_frame.clear()
        
        for store in self.stores:
            self.checkbox_frame.add_checkbox(store['name'], store['name'])
    
    def find_stores(self, base_folder):
        """Находит все папки с торговыми точками в базовой папке"""
        stores = []
        
        for item in os.listdir(base_folder):
            item_path = os.path.join(base_folder, item)
            if os.path.isdir(item_path):
                # Ищем файлы данных с разными вариантами имен
                data_files = self.find_data_files(item_path)
                limits_files = self.find_limits_files(item_path)
                
                if data_files and limits_files:
                    # Берем первый найденный файл каждого типа
                    data_file = data_files[0]
                    limits_file = limits_files[0]
                    
                    stores.append({
                        'name': item,
                        'folder': item_path,
                        'data_file': data_file,
                        'limits_file': limits_file
                    })
        
        return stores
    
    def find_data_files(self, folder):
        """Находит файлы с данными в папке"""
        possible_names = [
            "Data.xlsx", "data.xlsx", "DATA.xlsx",
            "Data.xls", "data.xls", "DATA.xls"
        ]
        
        found_files = []
        for name in possible_names:
            file_path = os.path.join(folder, name)
            if os.path.exists(file_path):
                found_files.append(file_path)
        
        return found_files
    
    def find_limits_files(self, folder):
        """Находит файлы с лимитами в папке"""
        possible_names = [
            "limits.xlsx", "Limits.xlsx", "LIMITS.xlsx",
            "limits.xls", "Limits.xls", "LIMITS.xls"
        ]
        
        found_files = []
        for name in possible_names:
            file_path = os.path.join(folder, name)
            if os.path.exists(file_path):
                found_files.append(file_path)
        
        return found_files
    
    def diagnose_folder(self):
        """Выполняет диагностику выбранной папки"""
        if not self.base_folder:
            messagebox.showwarning("Диагностика", "Сначала выберите папку")
            return
        
        self.log_text.delete(1.0, tk.END)
        self.log_message("=== ДИАГНОСТИКА ПАПКИ ===")
        self.log_message(f"Папка: {self.base_folder}")
        
        items = os.listdir(self.base_folder)
        self.log_message(f"Объектов в папке: {len(items)}")
        
        for i, item in enumerate(items):
            item_path = os.path.join(self.base_folder, item)
            if os.path.isdir(item_path):
                self.log_message(f"\n[{i+1}] Папка: {item}")
                
                # Проверяем файлы в папке
                data_files = self.find_data_files(item_path)
                limits_files = self.find_limits_files(item_path)
                
                if data_files:
                    self.log_message(f"   ✓ Найдены файлы данных: {[os.path.basename(f) for f in data_files]}")
                else:
                    self.log_message(f"   ✗ Файлы данных не найдены")
                
                if limits_files:
                    self.log_message(f"   ✓ Найдены файлы лимитов: {[os.path.basename(f) for f in limits_files]}")
                else:
                    self.log_message(f"   ✗ Файлы лимитов не найдены")
                
                if data_files and limits_files:
                    self.log_message(f"   ✓ Папка готова к обработке")
                else:
                    self.log_message(f"   ✗ Папка НЕ готова к обработке")
        
        self.log_message("\n=== ДИАГНОСТИКА ЗАВЕРШЕНА ===")
    
    def open_add_limit_window(self):
        """Открывает окно для добавления лимитов"""
        selected_names = self.checkbox_frame.get_selected()
        if not selected_names:
            messagebox.showwarning("Внимание", "Сначала выберите торговые точки")
            return
        
        # Создаем окно ввода
        self.limit_window = tk.Toplevel(self.root)
        self.limit_window.title("Добавить лимиты")
        self.limit_window.geometry("500x300")
        self.limit_window.resizable(False, False)
        self.limit_window.grab_set()
        
        # Центрируем окно
        self.limit_window.update_idletasks()
        x = (self.limit_window.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.limit_window.winfo_screenheight() // 2) - (300 // 2)
        self.limit_window.geometry(f'500x300+{x}+{y}')
        
        # Инструкция
        instruction = tk.Label(self.limit_window, 
                              text="Введите лимиты в формате: Товар - лимит\n"
                                   "Пример: Дарксайд 25 - 2\n"
                                   "Можно вводить несколько лимитов, каждый с новой строки:",
                              font=('Arial', 9),
                              justify=tk.LEFT)
        instruction.pack(pady=10, padx=10, anchor='w')
        
        # Текстовое поле для ввода
        self.limit_text = tk.Text(self.limit_window, height=8, width=50, font=('Arial', 9))
        self.limit_text.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        self.limit_text.focus()
        
        # Фрейм для кнопок
        button_frame = tk.Frame(self.limit_window)
        button_frame.pack(pady=10)
        
        add_btn = tk.Button(button_frame, text="Добавить", command=self.add_limits,
                           font=('Arial', 10), bg='#4CAF50', fg='white', width=10)
        add_btn.pack(side=tk.LEFT, padx=5)
        
        cancel_btn = tk.Button(button_frame, text="Отмена", command=self.limit_window.destroy,
                              font=('Arial', 10), width=10)
        cancel_btn.pack(side=tk.LEFT, padx=5)
    
    def add_limits(self):
        """Добавляет лимиты в выбранные точки"""
        try:
            # Получаем введенный текст
            input_text = self.limit_text.get(1.0, tk.END).strip()
            if not input_text:
                messagebox.showwarning("Внимание", "Введите лимиты для добавления")
                return
            
            # Парсим ввод
            limits_to_add = self.parse_limits_input(input_text)
            if not limits_to_add:
                messagebox.showerror("Ошибка", "Неверный формат ввода. Используйте: Товар - лимит")
                return
            
            # Получаем выбранные точки
            selected_names = self.checkbox_frame.get_selected()
            selected_stores = [store for store in self.stores if store['name'] in selected_names]
            
            # Обрабатываем каждую точку
            success_count = 0
            self.log_message("\n=== ДОБАВЛЕНИЕ ЛИМИТОВ ===")
            
            for store in selected_stores:
                try:
                    if self.add_limits_to_store(store, limits_to_add):
                        success_count += 1
                        self.log_message(f"✓ {store['name']}: добавлено {len(limits_to_add)} лимитов")
                    else:
                        self.log_message(f"✗ {store['name']}: ошибка при добавлении")
                except Exception as e:
                    self.log_message(f"✗ {store['name']}: {str(e)}")
            
            self.log_message(f"Успешно обработано: {success_count}/{len(selected_stores)} точек")
            messagebox.showinfo("Готово", f"Лимиты добавлены в {success_count} точек")
            self.limit_window.destroy()
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении лимитов: {str(e)}")
    
    def parse_limits_input(self, input_text):
        """Парсит ввод пользователя в список кортежей (товар, лимит)"""
        limits = []
        lines = input_text.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # Разделяем по " - " 
            if ' - ' in line:
                parts = line.split(' - ', 1)
                if len(parts) == 2:
                    product = parts[0].strip()
                    try:
                        limit = int(parts[1].strip())
                        limits.append((product, limit))
                    except ValueError:
                        continue
        
        return limits
    
    def add_limits_to_store(self, store_info, limits_to_add):
        """Добавляет лимиты в конкретную торговую точку"""
        limits_file = store_info['limits_file']
        
        try:
            # Читаем весь файл, сохраняя все листы
            with pd.ExcelFile(limits_file) as xls:
                sheets_dict = {}
                for sheet_name in xls.sheet_names:
                    sheets_dict[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
            
            # Обновляем лист Limits
            if 'Limits' in sheets_dict:
                limits_df = sheets_dict['Limits']
            else:
                # Создаем новый лист если его нет
                limits_df = pd.DataFrame(columns=['Товар', 'Лимиты'])
            
            # Добавляем новые лимиты
            for product, limit in limits_to_add:
                # Проверяем, есть ли уже такой товар
                mask = limits_df['Товар'].astype(str) == product
                if mask.any():
                    # Обновляем существующий лимит
                    limits_df.loc[mask, 'Лимиты'] = limit
                else:
                    # Добавляем новую строку
                    new_row = pd.DataFrame({'Товар': [product], 'Лимиты': [limit]})
                    limits_df = pd.concat([limits_df, new_row], ignore_index=True)
            
            # Сохраняем обратно в файл
            with pd.ExcelWriter(limits_file, engine='openpyxl') as writer:
                for sheet_name, df in sheets_dict.items():
                    if sheet_name == 'Limits':
                        df = limits_df
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            return True
            
        except Exception as e:
            print(f"Ошибка при добавлении лимитов в {store_info['name']}: {str(e)}")
            return False
    
    def find_best_match(self, product_name, limits_dict):
        """
        Находит наиболее подходящий лимит для товара используя регулярные выражения.
        Все слова из лимита должны присутствовать в названии товара в любом порядке.
        """
        best_match = None
        best_match_score = 0
        
        for limit_item in limits_dict.keys():
            # Разбиваем лимит на отдельные слова (игнорируем пустые)
            words = [word.strip() for word in limit_item.split() if word.strip()]
            
            if not words:
                continue
                
            # Создаем паттерн для проверки что ВСЕ слова присутствуют
            # Используем lookahead assertions для проверки всех слов в любом порядке
            patterns = []
            for word in words:
                # Экранируем специальные символы в словах
                escaped_word = re.escape(word)
                patterns.append(f"(?=.*{escaped_word})")
            
            # Комбинируем все паттерны
            full_pattern = "".join(patterns) + ".*"
            
            # Проверяем соответствие
            if re.search(full_pattern, product_name, re.IGNORECASE):
                # Оцениваем по количеству слов и длине совпадения
                score = len(words) * 1000 + len(limit_item)
                if score > best_match_score:
                    best_match = limit_item
                    best_match_score = score
        
        return best_match
    
    def reduce_order(self, df):
        """
        Сокращает заявку по заданным правилам:
        1. Удаляет строки, где заказ <= 1/3 от лимита
        2. Удаляет строки, где лимит = 2 и заказ = 1
        """
        initial_count = len(df)
        
        # Первая фильтрация: заказ <= 1/3 от лимита
        df = df[~(df['Заказ'] <= df['Лимиты'] / 3)]
        
        # Вторая фильтрация: лимит = 2 и заказ = 1
        df = df[~((df['Лимиты'] == 2) & (df['Заказ'] == 1))]
        
        removed_count = initial_count - len(df)
        return df, removed_count
    
    def process_store(self, store_info):
        """Обрабатывает одну торговую точку"""
        try:
            data_file = store_info['data_file']
            limits_file = store_info['limits_file']
            store_name = store_info['name']
            output_file = os.path.join(self.base_folder, f"{store_name}.xlsx")
            
            # Загрузка данных
            df = pd.read_excel(data_file)
            limits_df = pd.read_excel(limits_file, sheet_name="Limits")
            
            # Обрабатываем пустые значения в столбце "Остаток"
            df['Остаток'] = pd.to_numeric(df['Остаток'], errors='coerce').fillna(0)
            
            # Создаем словарь лимитов
            limits = dict(zip(limits_df['Товар'].astype(str), limits_df['Лимиты']))
            
            # Преобразуем товары в строки
            df['Товар'] = df['Товар'].astype(str)

            # Определяем какие товары имеют лимиты И лимит > 0
            def should_keep(product_name):
                """Проверяет, нужно ли сохранять товар"""
                best_match = self.find_best_match(product_name, limits)
                # Сохраняем только если найден лимит И лимит > 0
                return best_match is not None and limits[best_match] > 0
            
            # Фильтруем - оставляем только товары с положительными лимитами
            df = df[df['Товар'].apply(should_keep)]
            
            # Функция для расчета заказа
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
            
            # ФИНАЛЬНАЯ ФИЛЬТРАЦИЯ: удаляем все строки, где заказ равен 0
            initial_zero_count = len(df[df['Заказ'] == 0])
            df = df[df['Заказ'] > 0]
            
            # Применяем сокращение заявки, если включена опция
            reduced_count = 0
            if self.reduce_order_var.get():
                df, reduced_count = self.reduce_order(df)
                self.log_message(f"  Сокращено позиций: {reduced_count}")
            
            # Сохраняем результат
            df.to_excel(output_file, index=False)
            
            return True, f"✓ {store_name} - Успешно (осталось {len(df)} позиций)"
            
        except Exception as e:
            return False, f"✗ {store_name} - Ошибка: {str(e)}"
    
    def log_message(self, message):
        """Добавляет сообщение в лог"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def process_selected_stores(self):
        """Обрабатывает только выбранные торговые точки"""
        if not self.base_folder:
            messagebox.showerror("Ошибка", "Сначала выберите папку с автозаявками")
            return
        
        selected_names = self.checkbox_frame.get_selected()
        if not selected_names:
            messagebox.showerror("Ошибка", "Выберите хотя бы одну торговую точку")
            return
        
        # Фильтруем магазины по выбранным именам
        selected_stores = [store for store in self.stores if store['name'] in selected_names]
        
        # Настройка интерфейса для обработки
        self.process_btn.config(state='disabled', bg='gray', text="Обработка...")
        self.log_text.delete(1.0, tk.END)
        self.progress['value'] = 0
        self.progress['maximum'] = len(selected_stores)
        
        # Добавляем информацию о режиме обработки
        if self.reduce_order_var.get():
            self.log_message("=== РЕЖИМ СОКРАЩЕНИЯ ЗАЯВКИ ===")
            self.log_message("Будут удалены:")
            self.log_message("  - Позиции с заказом ≤ 1/3 от лимита")
            self.log_message("  - Позиции с лимитом=2 и заказом=1")
            self.log_message("")
        
        # Обработка каждой выбранной торговой точки
        success_count = 0
        for i, store_info in enumerate(selected_stores):
            store_name = store_info['name']
            self.progress_label.config(text=f"Обработка: {store_name} ({i+1}/{len(selected_stores)})")
            self.progress['value'] = i
            
            success, message = self.process_store(store_info)
            
            self.log_message(message)
            if success:
                success_count += 1
            
            self.root.update()
        
        # Завершение
        self.progress['value'] = len(selected_stores)
        self.progress_label.config(text="Обработка завершена")
        self.process_btn.config(state='normal', bg='#4CAF50', text="Обработать выбранные точки")
        
        messagebox.showinfo("Завершено", 
                          f"Обработка завершена!\n\n"
                          f"Успешно: {success_count} из {len(selected_stores)}")

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