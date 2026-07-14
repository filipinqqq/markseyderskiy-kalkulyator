"""
МАРКШЕЙДЕРСКИЙ КАЛЬКУЛЯТОР
Версия 2.0 - Полностью объектно-ориентированная архитектура
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
import math
import os
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass


# ============================================================
# МОДЕЛЬ (Model)
# ============================================================

@dataclass
class WellCoordinates:
    number: float
    x: float
    y: float
    
    def __post_init__(self):
        if not isinstance(self.number, (int, float)):
            raise ValueError(f"Номер должен быть числом, получено: {type(self.number)}")
        if not isinstance(self.x, (int, float)):
            raise ValueError(f"X должен быть числом, получено: {type(self.x)}")
        if not isinstance(self.y, (int, float)):
            raise ValueError(f"Y должен быть числом, получено: {type(self.y)}")
    
    def to_tuple(self) -> Tuple[float, float, float]:
        return (self.number, self.x, self.y)


@dataclass
class ComparisonResult:
    well_number: float
    project_x: float
    project_y: float
    fact_x: float
    fact_y: float
    deviation: float
    
    @property
    def dx(self) -> float:
        return self.project_x - self.fact_x
    
    @property
    def dy(self) -> float:
        return self.project_y - self.fact_y
    
    def to_list(self) -> List:
        return [
            f"{self.well_number:.0f}",
            f"{self.project_x:.3f}",
            f"{self.project_y:.3f}",
            f"{self.fact_x:.3f}",
            f"{self.fact_y:.3f}",
            f"{self.deviation:.3f}"
        ]


class ExcelReader:
    def __init__(self):
        self._last_error: Optional[str] = None
    
    def read_coordinates(self, file_path: str) -> List[WellCoordinates]:
        self._last_error = None
        
        try:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
            worksheet = workbook.active
            coordinates = []
            
            for row in worksheet.iter_rows(values_only=True):
                if not row or row[0] is None:
                    continue
                
                try:
                    number = float(row[0])
                    x = float(row[1]) if row[1] is not None else None
                    y = float(row[2]) if row[2] is not None else None
                    
                    if x is not None and y is not None:
                        coordinates.append(WellCoordinates(number, x, y))
                except (ValueError, TypeError):
                    continue
            
            return coordinates
            
        except Exception as e:
            self._last_error = str(e)
            return []
    
    def get_last_error(self) -> Optional[str]:
        return self._last_error
    
    def validate_file(self, file_path: str) -> bool:
        if not os.path.exists(file_path):
            self._last_error = f"Файл не найден: {file_path}"
            return False
        
        if not file_path.endswith(('.xlsx', '.xls')):
            self._last_error = "Неверный формат файла. Нужен .xlsx или .xls"
            return False
        
        return True


class Calculator:
    @staticmethod
    def calculate_deviation(project: WellCoordinates, fact: WellCoordinates) -> float:
        dx = project.x - fact.x
        dy = project.y - fact.y
        return math.sqrt(dx**2 + dy**2)
    
    @staticmethod
    def compare_wells(project_list: List[WellCoordinates], 
                     fact_list: List[WellCoordinates]) -> List[ComparisonResult]:
        project_dict = {well.number: well for well in project_list}
        fact_dict = {well.number: well for well in fact_list}
        
        common_numbers = set(project_dict.keys()) & set(fact_dict.keys())
        
        results = []
        for number in common_numbers:
            project = project_dict[number]
            fact = fact_dict[number]
            
            deviation = Calculator.calculate_deviation(project, fact)
            
            results.append(ComparisonResult(
                well_number=number,
                project_x=project.x,
                project_y=project.y,
                fact_x=fact.x,
                fact_y=fact.y,
                deviation=deviation
            ))
        
        return sorted(results, key=lambda x: x.deviation, reverse=True)
    
    @staticmethod
    def calculate_statistics(results: List[ComparisonResult]) -> Dict[str, float]:
        if not results:
            return {
                'count': 0,
                'average': 0.0,
                'max': 0.0,
                'min': 0.0,
                'total': 0.0
            }
        
        deviations = [r.deviation for r in results]
        
        return {
            'count': len(results),
            'average': sum(deviations) / len(deviations),
            'max': max(deviations),
            'min': min(deviations),
            'total': sum(deviations)
        }


class DataManager:
    _instance: Optional['DataManager'] = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if not hasattr(self, '_initialized'):
            self._project_data: List[WellCoordinates] = []
            self._fact_data: List[WellCoordinates] = []
            self._results: List[ComparisonResult] = []
            self._project_file: Optional[str] = None
            self._fact_file: Optional[str] = None
            self._initialized = True
    
    def set_project_data(self, data: List[WellCoordinates], filename: str):
        self._project_data = data
        self._project_file = filename
    
    def set_fact_data(self, data: List[WellCoordinates], filename: str):
        self._fact_data = data
        self._fact_file = filename
    
    def set_results(self, results: List[ComparisonResult]):
        self._results = results
    
    def get_project_data(self) -> List[WellCoordinates]:
        return self._project_data
    
    def get_fact_data(self) -> List[WellCoordinates]:
        return self._fact_data
    
    def get_results(self) -> List[ComparisonResult]:
        return self._results
    
    def get_project_file(self) -> Optional[str]:
        return self._project_file
    
    def get_fact_file(self) -> Optional[str]:
        return self._fact_file
    
    def has_data(self) -> bool:
        return bool(self._project_data) and bool(self._fact_data)
    
    def clear_data(self):
        self._project_data = []
        self._fact_data = []
        self._results = []
        self._project_file = None
        self._fact_file = None


# ============================================================
# ПРЕДСТАВЛЕНИЕ (View)
# ============================================================

class MainWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Маркшейдерский калькулятор")
        self.root.geometry("900x600")
        
        self.data_manager = DataManager()
        self.excel_reader = ExcelReader()
        self.calculator = Calculator()
        
        self._create_header()
        self._create_control_panel()
        self._create_status_bar()
        self._create_info_panel()
        self._create_result_table()
        self._create_stats_panel()
        
        self._update_status("Загрузите файлы проекта и факта", "blue")
    
    def _create_header(self):
        header_frame = tk.Frame(self.root)
        header_frame.pack(pady=10)
        
        tk.Label(header_frame, text="🏔️ Горный университет - Маркшейдерия", 
                font=("Arial", 18, "bold")).pack()
        tk.Label(header_frame, text="Сравнение проектных и фактических координат скважин", 
                font=("Arial", 12)).pack()
    
    def _create_control_panel(self):
        panel = tk.Frame(self.root)
        panel.pack(pady=10)
        
        self.btn_project = tk.Button(panel, text="📐 Загрузить ПРОЕКТ",
                                    command=self._on_load_project,
                                    width=20, height=2,
                                    bg="#4CAF50", fg="white",
                                    font=("Arial", 10, "bold"))
        self.btn_project.grid(row=0, column=0, padx=5)
        
        self.btn_fact = tk.Button(panel, text="📊 Загрузить ФАКТ",
                                 command=self._on_load_fact,
                                 width=20, height=2,
                                 bg="#2196F3", fg="white",
                                 font=("Arial", 10, "bold"))
        self.btn_fact.grid(row=0, column=1, padx=5)
        
        self.btn_calculate = tk.Button(panel, text="🚀 Рассчитать",
                                      command=self._on_calculate,
                                      width=20, height=2,
                                      bg="#FF9800", fg="white",
                                      font=("Arial", 10, "bold"))
        self.btn_calculate.grid(row=0, column=2, padx=5)
        
        self.btn_clear = tk.Button(panel, text="🗑️ Очистить",
                                  command=self._on_clear,
                                  width=20, height=2,
                                  bg="#f44336", fg="white",
                                  font=("Arial", 10, "bold"))
        self.btn_clear.grid(row=0, column=3, padx=5)
    
    def _create_status_bar(self):
        self.status_label = tk.Label(self.root, text="Готов к работе",
                                    font=("Arial", 10), fg="blue")
        self.status_label.pack(pady=5)
    
    def _create_info_panel(self):
        info_frame = tk.Frame(self.root)
        info_frame.pack(pady=5, fill=tk.X, padx=20)
        
        self.project_info = tk.Label(info_frame, text="Проект: не загружен",
                                    font=("Arial", 10), anchor="w")
        self.project_info.pack(side=tk.LEFT, padx=10)
        
        self.fact_info = tk.Label(info_frame, text="Факт: не загружен",
                                 font=("Arial", 10), anchor="w")
        self.fact_info.pack(side=tk.LEFT, padx=10)
    
    def _create_result_table(self):
        result_frame = tk.Frame(self.root)
        result_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        
        scrollbar = tk.Scrollbar(result_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        columns = ("Номер", "X_проект", "Y_проект", "X_факт", "Y_факт", "Отклонение")
        
        self.tree = ttk.Treeview(result_frame, yscrollcommand=scrollbar.set,
                                columns=columns, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.tree.yview)
        
        headers = {
            "Номер": "Номер скважины",
            "X_проект": "X проект",
            "Y_проект": "Y проект",
            "X_факт": "X факт",
            "Y_факт": "Y факт",
            "Отклонение": "Отклонение (м)"
        }
        
        for col, text in headers.items():
            self.tree.heading(col, text=text)
            self.tree.column(col, width=100, anchor="center")
    
    def _create_stats_panel(self):
        self.stats_label = tk.Label(self.root, text="", font=("Arial", 10), fg="green")
        self.stats_label.pack(pady=5)
    
    def _on_load_project(self):
        file_path = filedialog.askopenfilename(
            title="Выберите файл проекта",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return
        
        if not self.excel_reader.validate_file(file_path):
            messagebox.showerror("Ошибка", self.excel_reader.get_last_error())
            return
        
        data = self.excel_reader.read_coordinates(file_path)
        if not data:
            messagebox.showerror("Ошибка", f"Не удалось прочитать файл:\n{self.excel_reader.get_last_error()}")
            return
        
        self.data_manager.set_project_data(data, os.path.basename(file_path))
        self.project_info.config(text=f"Проект: {os.path.basename(file_path)} ({len(data)} скважин)")
        self._update_status(f"✅ Проект загружен: {len(data)} скважин", "green")
    
    def _on_load_fact(self):
        file_path = filedialog.askopenfilename(
            title="Выберите файл факта",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return
        
        if not self.excel_reader.validate_file(file_path):
            messagebox.showerror("Ошибка", self.excel_reader.get_last_error())
            return
        
        data = self.excel_reader.read_coordinates(file_path)
        if not data:
            messagebox.showerror("Ошибка", f"Не удалось прочитать файл:\n{self.excel_reader.get_last_error()}")
            return
        
        self.data_manager.set_fact_data(data, os.path.basename(file_path))
        self.fact_info.config(text=f"Факт: {os.path.basename(file_path)} ({len(data)} скважин)")
        self._update_status(f"✅ Факт загружен: {len(data)} скважин", "green")
    
    def _on_calculate(self):
        self._clear_table()
        
        if not self.data_manager.get_project_data():
            messagebox.showwarning("Внимание", "Сначала загрузите файл ПРОЕКТА!")
            return
        
        if not self.data_manager.get_fact_data():
            messagebox.showwarning("Внимание", "Сначала загрузите файл ФАКТА!")
            return
        
        project_data = self.data_manager.get_project_data()
        fact_data = self.data_manager.get_fact_data()
        
        results = self.calculator.compare_wells(project_data, fact_data)
        
        if not results:
            messagebox.showinfo("Результат", "Нет совпадающих номеров скважин!")
            return
        
        self.data_manager.set_results(results)
        
        for result in results:
            self.tree.insert("", "end", values=result.to_list())
        
        stats = self.calculator.calculate_statistics(results)
        self._update_statistics(stats)
        self._update_status(f"✅ Рассчитано {stats['count']} скважин", "green")
    
    def _on_clear(self):
        self._clear_table()
        self.data_manager.clear_data()
        self.project_info.config(text="Проект: не загружен")
        self.fact_info.config(text="Факт: не загружен")
        self.stats_label.config(text="")
        self._update_status("Данные очищены", "blue")
    
    def _clear_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
    
    def _update_status(self, text: str, color: str = "blue"):
        self.status_label.config(text=text, fg=color)
    
    def _update_statistics(self, stats: Dict[str, float]):
        if stats['count'] == 0:
            self.stats_label.config(text="Нет данных для статистики")
            return
        
        self.stats_label.config(
            text=f"📊 Всего: {stats['count']} скважин | "
                 f"Среднее: {stats['average']:.3f} м | "
                 f"Макс: {stats['max']:.3f} м | "
                 f"Мин: {stats['min']:.3f} м"
        )


# ============================================================
# ТОЧКА ВХОДА
# ============================================================

class Application:
    @staticmethod
    def run():
        root = tk.Tk()
        app = MainWindow(root)
        root.mainloop()


if __name__ == "__main__":
    Application.run()