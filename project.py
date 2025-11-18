import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import pandas as pd
from datetime import datetime
import sqlite3
import hashlib
import time
import json


class DatabaseManager:
    
    def __init__(self, db_name="ceramics.db"):
        self.db_name = db_name
        self.conn = None
        self.create_connection()
        self.create_tables()
        self.init_default_data()
    
    def create_connection(self):
        try:
            self.conn = sqlite3.connect(self.db_name)
            self.conn.row_factory = sqlite3.Row
            print(f"Подключено к БД: {self.db_name}")
        except sqlite3.Error as e:
            print(f"Ошибка подключения: {e}")
    
    def create_tables(self):
        if self.conn is None:
            return
        
        cursor = self.conn.cursor()
        
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                user_id INTEGER PRIMARY KEY AUTOINCREMENT,
                login TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL CHECK(role IN ('researcher', 'admin')),
                created_date DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS materials (
                material_id INTEGER PRIMARY KEY AUTOINCREMENT,
                material_name TEXT UNIQUE NOT NULL,
                material_type TEXT,
                description TEXT,
                created_date DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS model_coefficients (
                coefficient_id INTEGER PRIMARY KEY AUTOINCREMENT,
                material_id INTEGER NOT NULL,
                a0 REAL NOT NULL,
                a1 REAL NOT NULL,
                a2 REAL NOT NULL,
                a3 REAL NOT NULL,
                a4 REAL NOT NULL,
                a5 REAL NOT NULL,
                valid_from DATE,
                valid_to DATE,
                comment TEXT,
                created_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(material_id) REFERENCES materials(material_id)
            )
        """)
        
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS calculation_sessions (
                session_id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                material_id INTEGER NOT NULL,
                pg_min REAL,
                pg_max REAL,
                pg_step REAL,
                temp_min INTEGER,
                temp_max INTEGER,
                temp_step INTEGER,
                num_points INTEGER,
                operations_count INTEGER,
                exec_time_sec REAL,
                result_summary TEXT,
                created_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(user_id) REFERENCES users(user_id),
                FOREIGN KEY(material_id) REFERENCES materials(material_id)
            )
        """)
        
        self.conn.commit()
        print("Таблицы созданы")
    
    def init_default_data(self):
        cursor = self.conn.cursor()
        
        try:
            cursor.execute("SELECT COUNT(*) as cnt FROM users")
            if cursor.fetchone()['cnt'] == 0:
                researcher_hash = hashlib.sha256("pass123".encode()).hexdigest()
                admin_hash = hashlib.sha256("admin123".encode()).hexdigest()
                
                cursor.execute(
                    "INSERT INTO users (login, password_hash, role) VALUES (?, ?, ?)",
                    ("researcher", researcher_hash, "researcher")
                )
                cursor.execute(
                    "INSERT INTO users (login, password_hash, role) VALUES (?, ?, ?)",
                    ("admin", admin_hash, "admin")
                )
            
            cursor.execute("SELECT COUNT(*) as cnt FROM materials")
            if cursor.fetchone()['cnt'] == 0:
                cursor.execute(
                    """INSERT INTO materials (material_name, material_type, description) 
                       VALUES (?, ?, ?)""",
                    ("Карбид вольфрама-никель", "Твёрдый сплав", 
                     "WC-Ni композит для производства режущего инструмента")
                )
                material_id = cursor.lastrowid
                
                cursor.execute(
                    """INSERT INTO model_coefficients 
                       (material_id, a0, a1, a2, a3, a4, a5, valid_from) 
                       VALUES (?, ?, ?, ?, ?, ?, ?, DATE('now'))""",
                    (material_id, -17.46, -0.00622, 0.04293, 1.5e-5, -1.4e-5, -5e-9)
                )
            
            self.conn.commit()
            
        except sqlite3.Error as e:
            print(f"Ошибка инициализации: {e}")
    
    def verify_user(self, login, password):
        cursor = self.conn.cursor()
        password_hash = hashlib.sha256(password.encode()).hexdigest()
        cursor.execute(
            "SELECT user_id, role FROM users WHERE login = ? AND password_hash = ?",
            (login, password_hash)
        )
        result = cursor.fetchone()
        if result:
            return True, dict(result)
        return False, None
    
    def get_materials(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT material_id, material_name FROM materials ORDER BY material_name")
        return cursor.fetchall()
    
    def add_material(self, material_name, material_type, description, coeffs):
        cursor = self.conn.cursor()
        try:
            cursor.execute(
                "INSERT INTO materials (material_name, material_type, description) VALUES (?, ?, ?)",
                (material_name, material_type, description)
            )
            material_id = cursor.lastrowid
            
            cursor.execute(
                """INSERT INTO model_coefficients 
                   (material_id, a0, a1, a2, a3, a4, a5, valid_from) 
                   VALUES (?, ?, ?, ?, ?, ?, ?, DATE('now'))""",
                (material_id, coeffs['a0'], coeffs['a1'], coeffs['a2'], 
                 coeffs['a3'], coeffs['a4'], coeffs['a5'])
            )
            
            self.conn.commit()
            return material_id
            
        except sqlite3.IntegrityError:
            raise ValueError("Материал с таким названием уже существует")
    
    def get_coefficients(self, material_id):
        cursor = self.conn.cursor()
        cursor.execute(
            """SELECT a0, a1, a2, a3, a4, a5 FROM model_coefficients 
               WHERE material_id = ? ORDER BY created_date DESC LIMIT 1""",
            (material_id,)
        )
        result = cursor.fetchone()
        if result:
            return dict(result)
        return None
    
    def update_coefficients(self, material_id, coeffs):
        cursor = self.conn.cursor()
        cursor.execute(
            """INSERT INTO model_coefficients 
               (material_id, a0, a1, a2, a3, a4, a5, valid_from) 
               VALUES (?, ?, ?, ?, ?, ?, ?, DATE('now'))""",
            (material_id, coeffs['a0'], coeffs['a1'], coeffs['a2'], 
             coeffs['a3'], coeffs['a4'], coeffs['a5'])
        )
        self.conn.commit()
    
    def save_calculation_session(self, user_id, material_id, pg_min, pg_max, pg_step, 
                                 t_min, t_max, t_step, results_df, exec_time, operations):
        cursor = self.conn.cursor()
        
        result_summary = {
            "num_points": len(results_df),
            "min_density": float(results_df['rho'].min()),
            "max_density": float(results_df['rho'].max()),
            "mean_density": float(results_df['rho'].mean()),
            "std_density": float(results_df['rho'].std())
        }
        
        cursor.execute(
            """INSERT INTO calculation_sessions 
               (user_id, material_id, pg_min, pg_max, pg_step, temp_min, temp_max, temp_step, 
                num_points, operations_count, exec_time_sec, result_summary)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (user_id, material_id, pg_min, pg_max, pg_step, t_min, t_max, t_step, 
             len(results_df), operations, exec_time, json.dumps(result_summary))
        )
        
        self.conn.commit()


class CeramicsDensityApp:
    
    def __init__(self, root):
        self.root = root
        self.root.title("Исследование спекания керамических материалов")
        self.root.geometry("1600x1000")
        
        self.db = DatabaseManager()
        
        self.current_user = None
        self.current_user_id = None
        self.current_role = None
        self.current_data = None
        self.canvas_widget = None
        
        self.show_login_screen()
    
    def show_login_screen(self):
        self.clear_window()
        
        frame = ttk.Frame(self.root, padding="20")
        frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        ttk.Label(frame, text="ВХОД В СИСТЕМУ", font=("Arial", 16, "bold")).pack(pady=20)
        ttk.Label(frame, text="Логин:").pack()
        login_var = tk.StringVar()
        ttk.Entry(frame, textvariable=login_var, width=30).pack(pady=5)
        
        ttk.Label(frame, text="Пароль:").pack()
        password_var = tk.StringVar()
        ttk.Entry(frame, textvariable=password_var, width=30, show="*").pack(pady=5)
        
        def login():
            login = login_var.get()
            password = password_var.get()
            
            valid, user_data = self.db.verify_user(login, password)
            
            if valid:
                self.current_user = login
                self.current_user_id = user_data['user_id']
                self.current_role = user_data['role']
                
                if self.current_role == "admin":
                    self.show_admin_menu()
                else:
                    self.show_researcher_menu()
            else:
                messagebox.showerror("Ошибка", "Неверный логин или пароль!")
        
        ttk.Button(frame, text="Вход", command=login).pack(pady=20)
        ttk.Button(frame, text="Выход", command=self.root.quit).pack()
    
    def show_researcher_menu(self):
        self.clear_window()
        
        frame = ttk.Frame(self.root, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text=f"Добро пожаловать, {self.current_user}!", 
                 font=("Arial", 14, "bold")).pack(pady=10)
        
        ttk.Button(frame, text="Начать исследование", 
                  command=self.show_research_interface, width=30).pack(pady=10)
        ttk.Button(frame, text="Выход", command=self.show_login_screen, width=30).pack(pady=5)
    
    def show_admin_menu(self):
        self.clear_window()
        
        frame = ttk.Frame(self.root, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text=f"Администратор: {self.current_user}", 
                 font=("Arial", 14, "bold")).pack(pady=10)
        ttk.Button(frame, text="Управление коэффициентами", 
                  command=self.show_coefficients_editor, width=40).pack(pady=10)
        ttk.Button(frame, text="Добавить материал", 
                  command=self.show_add_material, width=40).pack(pady=10)
        ttk.Button(frame, text="Выход", command=self.show_login_screen, width=40).pack(pady=5)
    
    def show_add_material(self):
        self.clear_window()
        
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Добавление материала", 
                 font=("Arial", 14, "bold")).pack(pady=10)
        
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind("<Configure>", 
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        ttk.Label(scrollable_frame, text="Название:").pack()
        name_var = tk.StringVar()
        ttk.Entry(scrollable_frame, textvariable=name_var, width=40).pack(pady=5)
        
        ttk.Label(scrollable_frame, text="Тип:").pack()
        type_var = tk.StringVar()
        ttk.Entry(scrollable_frame, textvariable=type_var, width=40).pack(pady=5)
        
        ttk.Label(scrollable_frame, text="Описание:").pack()
        desc_var = tk.StringVar()
        ttk.Entry(scrollable_frame, textvariable=desc_var, width=40).pack(pady=5)
        
        coeff_frame = ttk.LabelFrame(scrollable_frame, text="Коэффициенты", padding="5")
        coeff_frame.pack(pady=10, fill=tk.BOTH)
        
        coeff_vars = {}
        for key in ["a0", "a1", "a2", "a3", "a4", "a5"]:
            ttk.Label(coeff_frame, text=f"{key}:").pack()
            var = tk.DoubleVar(value=0.0)
            ttk.Entry(coeff_frame, textvariable=var, width=20).pack(pady=2)
            coeff_vars[key] = var
        
        def save_material():
            if not name_var.get():
                messagebox.showerror("Ошибка", "Введите название!")
                return
            
            coeffs = {key: coeff_vars[key].get() for key in coeff_vars}
            
            try:
                self.db.add_material(name_var.get(), type_var.get(), desc_var.get(), coeffs)
                messagebox.showinfo("Успех", "Материал добавлен!")
                self.show_admin_menu()
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))
        
        ttk.Button(scrollable_frame, text="Сохранить", command=save_material).pack(pady=10)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        ttk.Button(main_frame, text="Назад", command=self.show_admin_menu).pack(pady=10)
    
    def show_coefficients_editor(self):
        self.clear_window()
        
        frame = ttk.Frame(self.root, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Редактирование коэффициентов", 
                 font=("Arial", 14, "bold")).pack(pady=10)
        
        materials = self.db.get_materials()
        material_names = [m['material_name'] for m in materials]
        material_ids = {m['material_name']: m['material_id'] for m in materials}
        
        material_var = tk.StringVar(value=material_names[0] if material_names else "")
        ttk.Label(frame, text="Материал:").pack()
        material_combo = ttk.Combobox(frame, textvariable=material_var, 
                                     values=material_names, state="readonly", width=30)
        material_combo.pack(pady=5)
        
        coeff_frame = ttk.LabelFrame(frame, text="Коэффициенты", padding="10")
        coeff_frame.pack(pady=20, fill=tk.BOTH, expand=True)
        
        coeff_vars = {}
        
        def load_coefficients(material_name):
            material_id = material_ids[material_name]
            coeffs = self.db.get_coefficients(material_id)
            if coeffs:
                for key in ["a0", "a1", "a2", "a3", "a4", "a5"]:
                    coeff_vars[key].set(coeffs[key])
        
        def on_material_change(event):
            load_coefficients(material_var.get())
        
        material_combo.bind("<<ComboboxSelected>>", on_material_change)
        
        for key in ["a0", "a1", "a2", "a3", "a4", "a5"]:
            ttk.Label(coeff_frame, text=f"{key}:").pack()
            var = tk.DoubleVar(value=0.0)
            ttk.Entry(coeff_frame, textvariable=var, width=20).pack(pady=3)
            coeff_vars[key] = var
        
        if material_names:
            load_coefficients(material_names[0])
        
        def save_coefficients():
            material_name = material_var.get()
            material_id = material_ids[material_name]
            coeffs = {key: coeff_vars[key].get() for key in coeff_vars}
            self.db.update_coefficients(material_id, coeffs)
            messagebox.showinfo("Успех", "Коэффициенты обновлены!")
        
        ttk.Button(frame, text="Сохранить", command=save_coefficients).pack(pady=10)
        ttk.Button(frame, text="Назад", command=self.show_admin_menu).pack()
    
    def show_research_interface(self):
        self.clear_window()
        
        top_frame = ttk.Frame(self.root)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)
        
        ttk.Label(top_frame, 
                 text="ИССЛЕДОВАНИЕ ВЛИЯНИЯ ДАВЛЕНИЯ И ТЕМПЕРАТУРЫ НА ПЛОТНОСТЬ",
                 font=("Arial", 12, "bold")).pack()
        
        left_frame = ttk.LabelFrame(self.root, text="Параметры исследования", padding="10")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=10, pady=10)
        
        materials = self.db.get_materials()
        material_names = [m['material_name'] for m in materials]
        material_ids_dict = {m['material_name']: m['material_id'] for m in materials}
        
        ttk.Label(left_frame, text="Материал:").pack()
        material_var = tk.StringVar(value=material_names[0] if material_names else "")
        ttk.Combobox(left_frame, textvariable=material_var, 
                    values=material_names, state="readonly").pack(pady=5)
        
        pg_min_var = tk.DoubleVar(value=40)
        pg_max_var = tk.DoubleVar(value=80)
        pg_step_var = tk.DoubleVar(value=2)
        t_min_var = tk.IntVar(value=1300)
        t_max_var = tk.IntVar(value=1500)
        t_step_var = tk.IntVar(value=10)
        
        ttk.Label(left_frame, text="Давление газа (атм):").pack(pady=(15, 0))
        ttk.Label(left_frame, text="мин:").pack()
        ttk.Spinbox(left_frame, from_=0, to=100, textvariable=pg_min_var, format="%.2f").pack()
        
        ttk.Label(left_frame, text="макс:").pack()
        ttk.Spinbox(left_frame, from_=0, to=100, textvariable=pg_max_var, format="%.2f").pack()
        
        ttk.Label(left_frame, text="шаг:").pack()
        ttk.Spinbox(left_frame, from_=0.1, to=10, textvariable=pg_step_var, format="%.2f").pack(pady=5)
        
        ttk.Label(left_frame, text="Температура (°C):").pack(pady=(15, 0))
        ttk.Label(left_frame, text="мин:").pack()
        ttk.Spinbox(left_frame, from_=1000, to=2000, textvariable=t_min_var).pack()
        
        ttk.Label(left_frame, text="макс:").pack()
        ttk.Spinbox(left_frame, from_=1000, to=2000, textvariable=t_max_var).pack()
        
        ttk.Label(left_frame, text="шаг:").pack()
        ttk.Spinbox(left_frame, from_=1, to=50, textvariable=t_step_var).pack(pady=5)
        
        right_frame = ttk.LabelFrame(self.root, text="Результаты", padding="10")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.tree = ttk.Treeview(right_frame, columns=["Pg", "T", "ρ"], height=8)
        self.tree.column("#0", width=0, stretch=tk.NO)
        self.tree.column("Pg", anchor=tk.CENTER, width=70)
        self.tree.column("T", anchor=tk.CENTER, width=70)
        self.tree.column("ρ", anchor=tk.CENTER, width=90)
        self.tree.heading("Pg", text="Pg (атм)")
        self.tree.heading("T", text="T (°C)")
        self.tree.heading("ρ", text="ρ (г/см³)")
        self.tree.pack(fill=tk.BOTH, expand=True, pady=10)
        
        stats_frame = ttk.LabelFrame(right_frame, text="Показатели экономичности", padding="10")
        stats_frame.pack(fill=tk.X, pady=5)
        
        self.stats_label = ttk.Label(stats_frame, text="", justify=tk.LEFT)
        self.stats_label.pack()
        
        calc_frame = ttk.LabelFrame(right_frame, text="Показатели расчётов", padding="10")
        calc_frame.pack(fill=tk.X, pady=5)
        
        self.calc_label = ttk.Label(calc_frame, text="", justify=tk.LEFT)
        self.calc_label.pack()
        
        button_frame = ttk.Frame(right_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        def calculate_and_display():
            try:
                material_name = material_var.get()
                material_id = material_ids_dict[material_name]
                pg_min = pg_min_var.get()
                pg_max = pg_max_var.get()
                pg_step = pg_step_var.get()
                t_min = t_min_var.get()
                t_max = t_max_var.get()
                t_step = t_step_var.get()
                
                if pg_min < 0 or pg_max < 0 or t_min < 0 or t_max < 0:
                    messagebox.showerror("Ошибка", "Параметры не могут быть отрицательными!")
                    return
                if pg_min >= pg_max or t_min >= t_max:
                    messagebox.showerror("Ошибка", "Минимум должен быть меньше максимума!")
                    return
                
                self.calculate_density(material_id, pg_min, pg_max, pg_step, 
                                      t_min, t_max, t_step, material_name, right_frame)
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")
        
        ttk.Button(button_frame, text="Рассчитать", command=calculate_and_display, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Сохранить (Excel)", 
                  command=lambda: self.save_report(material_var.get()), width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(left_frame, text="Выход", command=self.show_researcher_menu).pack(fill=tk.X, pady=5)
    
    def calculate_density(self, material_id, pg_min, pg_max, pg_step, t_min, t_max, t_step, material_name, parent_frame):
        start_time = time.time()
        
        coeffs_dict = self.db.get_coefficients(material_id)
        if not coeffs_dict:
            messagebox.showerror("Ошибка", "Коэффициенты не найдены!")
            return
        
        a0 = coeffs_dict['a0']
        a1 = coeffs_dict['a1']
        a2 = coeffs_dict['a2']
        a3 = coeffs_dict['a3']
        a4 = coeffs_dict['a4']
        a5 = coeffs_dict['a5']
        
        pg_values = np.arange(pg_min, pg_max + pg_step, pg_step)
        t_values = np.arange(t_min, t_max + t_step, t_step)
        
        data = []
        operations = 0
        ops_per_calc = 13
        
        for pg in pg_values:
            for t in t_values:
                rho = a0 + a1*pg + a2*t + a3*pg*t + a4*t**2 + a5*pg*t**2
                data.append({"Pg": pg, "T": t, "rho": rho})
                operations += ops_per_calc
        
        df = pd.DataFrame(data)
        self.current_data = df
        
        self.tree.delete(*self.tree.get_children())
        for idx, row in df.iterrows():
            self.tree.insert("", tk.END, values=[f"{row['Pg']:.2f}", f"{row['T']:.0f}", f"{row['rho']:.2f}"])
        
        self.plot_results(df, parent_frame)
        
        calc_time = time.time() - start_time
        
        stats_text = f"Время: {calc_time:.6f} с\nПамять: ~{len(df)*0.001:.2f} МБ"
        self.stats_label.config(text=stats_text)
        
        calc_text = f"Операции: {operations}\nМин ρ: {df['rho'].min():.2f}\nМакс ρ: {df['rho'].max():.2f}\nСредняя ρ: {df['rho'].mean():.2f}"
        self.calc_label.config(text=calc_text)
        
        self.db.save_calculation_session(self.current_user_id, material_id, pg_min, pg_max, pg_step,
                                        t_min, t_max, t_step, df, calc_time, operations)
    
    def plot_results(self, df, parent_frame):
        t_values_unique = sorted(df['T'].unique())
        pg_values_unique = sorted(df['Pg'].unique())
        
        t_indices = [0, len(t_values_unique)//2, -1]
        t_selected = [t_values_unique[i] for i in t_indices if i < len(t_values_unique)]
        
        pg_indices = [0, len(pg_values_unique)//2, -1]
        pg_selected = [pg_values_unique[i] for i in pg_indices if i < len(pg_values_unique)]
        
        fig = Figure(figsize=(10, 5), dpi=100)
        
        ax1 = fig.add_subplot(1, 2, 1)
        for t_val in t_selected:
            subset = df[df['T'] == t_val].sort_values('Pg')
            ax1.plot(subset['Pg'], subset['rho'], marker='o', linewidth=2, label=f'T={t_val}°C')
        ax1.set_xlabel('Давление газа Pg (атм)', fontsize=10)
        ax1.set_ylabel('Плотность ρ (г/см³)', fontsize=10)
        ax1.set_title('Зависимость плотности от давления', fontsize=11, fontweight='bold')
        ax1.legend(fontsize=9)
        ax1.grid(True, alpha=0.3)
        
        ax2 = fig.add_subplot(1, 2, 2)
        for pg_val in pg_selected:
            subset = df[df['Pg'] == pg_val].sort_values('T')
            ax2.plot(subset['T'], subset['rho'], marker='s', linewidth=2, label=f'Pg={pg_val:.2f} атм')
        ax2.set_xlabel('Температура T (°C)', fontsize=10)
        ax2.set_ylabel('Плотность ρ (г/см³)', fontsize=10)
        ax2.set_title('Зависимость плотности от температуры', fontsize=11, fontweight='bold')
        ax2.legend(fontsize=9)
        ax2.grid(True, alpha=0.3)
        
        fig.tight_layout()
        
        if self.canvas_widget:
            self.canvas_widget.get_tk_widget().destroy()
        
        self.canvas_widget = FigureCanvasTkAgg(fig, master=parent_frame)
        self.canvas_widget.draw()
        self.canvas_widget.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def save_report(self, material_name):
        if self.current_data is None:
            messagebox.showwarning("Внимание", "Сначала выполните расчёт!")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    self.current_data.to_excel(writer, sheet_name='Результаты', index=False)
                    
                    info_df = pd.DataFrame({
                        'Параметр': ['Материал', 'Дата', 'Мин ρ', 'Макс ρ', 'Средняя ρ'],
                        'Значение': [
                            material_name,
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            f"{self.current_data['rho'].min():.2f}",
                            f"{self.current_data['rho'].max():.2f}",
                            f"{self.current_data['rho'].mean():.2f}"
                        ]
                    })
                    info_df.to_excel(writer, sheet_name='Информация', index=False)
                
                messagebox.showinfo("Успех", f"Отчёт сохранён:\n{filename}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")
    
    def clear_window(self):
        for widget in self.root.winfo_children():
            widget.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = CeramicsDensityApp(root)
    root.mainloop()
