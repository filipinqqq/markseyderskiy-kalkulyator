import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
import math
import os

class МаркшейдерскийКалькулятор:
    def __init__(self, root):
        self.root = root
        self.root.title("Маркшейдерский калькулятор")
        self.root.geometry("900x600")
        
        self.project_data = None
        self.fact_data = None
        self.project_file = None
        self.fact_file = None
        
        # Заголовок
        tk.Label(root, text="🏔️ Горный университет - Маркшейдерия", 
                font=("Arial", 18, "bold")).pack(pady=10)
        tk.Label(root, text="Сравнение проектных и фактических координат скважин", 
                font=("Arial", 12)).pack(pady=5)
        
        # Рамка для кнопок загрузки
        frame_buttons = tk.Frame(root)
        frame_buttons.pack(pady=20)
        
        # Кнопка загрузки проекта
        self.btn_project = tk.Button(frame_buttons, text="📐 Загрузить ПРОЕКТ", 
                                    command=self.load_project, width=20, height=2,
                                    bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        self.btn_project.grid(row=0, column=0, padx=10)
        
        # Кнопка загрузки факта
        self.btn_fact = tk.Button(frame_buttons, text="📊 Загрузить ФАКТ", 
                                 command=self.load_fact, width=20, height=2,
                                 bg="#2196F3", fg="white", font=("Arial", 10, "bold"))
        self.btn_fact.grid(row=0, column=1, padx=10)
        
        # Кнопка расчета
        self.btn_calc = tk.Button(frame_buttons, text="🚀 Рассчитать", 
                                 command=self.calculate, width=20, height=2,
                                 bg="#FF9800", fg="white", font=("Arial", 10, "bold"))
        self.btn_calc.grid(row=0, column=2, padx=10)
        
        # Статус
        self.status_label = tk.Label(root, text="Загрузите файлы проекта и факта", 
                                    font=("Arial", 10), fg="blue")
        self.status_label.pack(pady=10)
        
        # Информация о файлах
        self.info_frame = tk.Frame(root)
        self.info_frame.pack(pady=10, fill=tk.X, padx=20)
        
        self.project_info = tk.Label(self.info_frame, text="Проект: не загружен", 
                                    font=("Arial", 10), anchor="w")
        self.project_info.pack(side=tk.LEFT, padx=10)
        
        self.fact_info = tk.Label(self.info_frame, text="Факт: не загружен", 
                                 font=("Arial", 10), anchor="w")
        self.fact_info.pack(side=tk.LEFT, padx=10)
        
        # Таблица результатов
        self.result_frame = tk.Frame(root)
        self.result_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        
        # Скролл для таблицы
        scrollbar = tk.Scrollbar(self.result_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree = ttk.Treeview(self.result_frame, yscrollcommand=scrollbar.set, 
                                 columns=("Номер", "X_проект", "Y_проект", "X_факт", "Y_факт", "Отклонение"))
        self.tree.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.tree.yview)
        
        # Заголовки таблицы
        self.tree.heading("#0", text="")
        self.tree.heading("Номер", text="Номер скважины")
        self.tree.heading("X_проект", text="X проект")
        self.tree.heading("Y_проект", text="Y проект")
        self.tree.heading("X_факт", text="X факт")
        self.tree.heading("Y_факт", text="Y факт")
        self.tree.heading("Отклонение", text="Отклонение (м)")
        
        # Ширина колонок
        self.tree.column("#0", width=0, stretch=False)
        self.tree.column("Номер", width=100)
        self.tree.column("X_проект", width=120)
        self.tree.column("Y_проект", width=120)
        self.tree.column("X_факт", width=120)
        self.tree.column("Y_факт", width=120)
        self.tree.column("Отклонение", width=120)
        
        # Статистика
        self.stats_label = tk.Label(root, text="", font=("Arial", 10), fg="green")
        self.stats_label.pack(pady=5)
    
    def load_project(self):
        file_path = filedialog.askopenfilename(
            title="Выберите файл проекта", 
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            data = self.read_excel(file_path)
            if data:
                self.project_data = data
                self.project_file = os.path.basename(file_path)
                self.project_info.config(text=f"Проект: {self.project_file} ({len(data)} скважин)")
                self.status_label.config(text=f"✅ Проект загружен: {len(data)} скважин", fg="green")
    
    def load_fact(self):
        file_path = filedialog.askopenfilename(
            title="Выберите файл факта", 
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            data = self.read_excel(file_path)
            if data:
                self.fact_data = data
                self.fact_file = os.path.basename(file_path)
                self.fact_info.config(text=f"Факт: {self.fact_file} ({len(data)} скважин)")
                self.status_label.config(text=f"✅ Факт загружен: {len(data)} скважин", fg="green")
    
    def read_excel(self, file_path):
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb.active
            
            data = []
            for row in ws.iter_rows(values_only=True):
                if row and row[0] is not None:
                    try:
                        # Берем первые 3 колонки: Номер, X, Y
                        num = float(row[0])
                        x = float(row[1]) if row[1] is not None else None
                        y = float(row[2]) if row[2] is not None else None
                        
                        if x is not None and y is not None:
                            data.append({
                                'Номер': num,
                                'X': x,
                                'Y': y
                            })
                    except:
                        continue
            
            return data if data else None
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось прочитать файл:\n{str(e)}")
            return None
    
    def calculate(self):
        # Очищаем таблицу
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if not self.project_data:
            messagebox.showwarning("Внимание", "Сначала загрузите файл ПРОЕКТА!")
            return
        
        if not self.fact_data:
            messagebox.showwarning("Внимание", "Сначала загрузите файл ФАКТА!")
            return
        
        # Создаем словари для поиска
        project_dict = {item['Номер']: item for item in self.project_data}
        fact_dict = {item['Номер']: item for item in self.fact_data}
        
        # Находим общие номера
        common = set(project_dict.keys()) & set(fact_dict.keys())
        
        if not common:
            messagebox.showinfo("Результат", "Нет совпадающих номеров скважин!")
            return
        
        # Считаем отклонения
        results = []
        for num in common:
            p = project_dict[num]
            f = fact_dict[num]
            
            dx = p['X'] - f['X']
            dy = p['Y'] - f['Y']
            deviation = math.sqrt(dx**2 + dy**2)
            
            results.append({
                'Номер': num,
                'X_проект': p['X'],
                'Y_проект': p['Y'],
                'X_факт': f['X'],
                'Y_факт': f['Y'],
                'Отклонение': deviation
            })
        
        # Сортируем по отклонению
        results = sorted(results, key=lambda x: x['Отклонение'], reverse=True)
        
        # Заполняем таблицу
        for r in results:
            self.tree.insert("", "end", values=(
                f"{r['Номер']:.0f}",
                f"{r['X_проект']:.3f}",
                f"{r['Y_проект']:.3f}",
                f"{r['X_факт']:.3f}",
                f"{r['Y_факт']:.3f}",
                f"{r['Отклонение']:.3f}"
            ))
        
        # Статистика
        deviations = [r['Отклонение'] for r in results]
        avg = sum(deviations) / len(deviations) if deviations else 0
        max_dev = max(deviations) if deviations else 0
        min_dev = min(deviations) if deviations else 0
        
        self.stats_label.config(
            text=f"📊 Всего: {len(results)} скважин | "
                 f"Среднее: {avg:.3f} м | "
                 f"Макс: {max_dev:.3f} м | "
                 f"Мин: {min_dev:.3f} м"
        )
        
        self.status_label.config(text=f"✅ Рассчитано {len(results)} скважин", fg="green")

# --- ЗАПУСК ---
if __name__ == "__main__":
    root = tk.Tk()
    app = МаркшейдерскийКалькулятор(root)
    root.mainloop()