import os
import pyodbc
import win32com.client as win32
from win32com.client import constants
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime, timedelta
import pythoncom
from tkcalendar import DateEntry
import shutil


# --- 1. Создание и резервное копирование БД ---
def create_access_database(db_path):
    if os.path.exists(db_path):
        try:
            os.remove(db_path)
        except PermissionError:
            messagebox.showerror("Ошибка", f"Файл {db_path} используется другим процессом.")
            return False

    try:
        pythoncom.CoInitialize()
        access_app = win32.Dispatch("Access.Application")
        access_app.NewCurrentDatabase(db_path)
        db = access_app.CurrentDb()

        # Таблица групп
        sql_groups = """
        CREATE TABLE Groups (
            group_id AUTOINCREMENT PRIMARY KEY,
            group_name TEXT,
            description MEMO
        );
        """
        db.Execute(sql_groups)

        # Таблица спортсменов
        sql_athletes = """
        CREATE TABLE Athletes (
            athlete_id AUTOINCREMENT PRIMARY KEY,
            name TEXT,
            birth_date DATETIME,
            phone TEXT,
            current_group_id LONG
        );
        """
        db.Execute(sql_athletes)

        # Таблица оплаты
        sql_payments = """
        CREATE TABLE Payments (
            payment_id AUTOINCREMENT PRIMARY KEY,
            athlete_id LONG,
            month_year TEXT,
            paid YESNO,
            CONSTRAINT NoDuplicatePayment UNIQUE (athlete_id, month_year)
        );
        """
        db.Execute(sql_payments)

        access_app.Quit()
        return True

    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось создать БД: {str(e)}")
        return False
    finally:
        pythoncom.CoUninitialize()


def backup_database(db_path):
    try:
        backup_dir = os.path.join(os.path.dirname(db_path), "backups")
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(backup_dir, f"backup_{timestamp}.accdb")
        shutil.copy2(db_path, backup_path)
        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось создать резервную копию: {str(e)}")
        return False


# --- 2. Подключение к БД ---
def connect_db():
    db_path = os.path.join(os.path.dirname(__file__), "sportclub.accdb")
    if not os.path.exists(db_path):
        if not create_access_database(db_path):
            return None

    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={db_path};'
    )
    try:
        return pyodbc.connect(conn_str)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось подключиться к БД: {str(e)}")
        return None


# --- 3. Функции работы с данными ---
def load_groups():
    conn = connect_db()
    if not conn:
        return []
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT group_id, group_name, description FROM Groups ORDER BY group_name")
        return cursor.fetchall()
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка загрузки групп: {str(e)}")
        return []
    finally:
        conn.close()


def load_athletes(group_id=None):
    conn = connect_db()
    if not conn:
        return []
    try:
        cursor = conn.cursor()
        if group_id:
            cursor.execute("""
                SELECT athlete_id, name, birth_date, phone 
                FROM Athletes 
                WHERE current_group_id = ? 
                ORDER BY name
            """, (group_id,))
        else:
            cursor.execute("SELECT athlete_id, name, birth_date, phone FROM Athletes ORDER BY name")
        return cursor.fetchall()
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка загрузки спортсменов: {str(e)}")
        return []
    finally:
        conn.close()


def add_group(name, description):
    conn = connect_db()
    if not conn:
        return False
    try:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO Groups (group_name, description) VALUES (?, ?)", (name, description))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка добавления группы: {str(e)}")
        return False
    finally:
        conn.close()


def update_group(group_id, name, description):
    conn = connect_db()
    if not conn:
        return False
    try:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE Groups 
            SET group_name = ?, description = ? 
            WHERE group_id = ?
        """, (name, description, group_id))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка обновления группы: {str(e)}")
        return False
    finally:
        conn.close()


def delete_group(group_id):
    conn = connect_db()
    if not conn:
        return False
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM Groups WHERE group_id = ?", (group_id,))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка удаления группы: {str(e)}")
        return False
    finally:
        conn.close()


def add_athlete(name, birth_date, phone, group_id):
    conn = connect_db()
    if not conn:
        return False
    try:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO Athletes (name, birth_date, phone, current_group_id) 
            VALUES (?, ?, ?, ?)
        """, (name, birth_date, phone, group_id))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка добавления спортсмена: {str(e)}")
        return False
    finally:
        conn.close()


def update_athlete(athlete_id, name, birth_date, phone, group_id):
    conn = connect_db()
    if not conn:
        return False
    try:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE Athletes 
            SET name = ?, birth_date = ?, phone = ?, current_group_id = ? 
            WHERE athlete_id = ?
        """, (name, birth_date, phone, group_id, athlete_id))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка обновления спортсмена: {str(e)}")
        return False
    finally:
        conn.close()


def delete_athlete(athlete_id):
    conn = connect_db()
    if not conn:
        return False
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM Athletes WHERE athlete_id = ?", (athlete_id,))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка удаления спортсмена: {str(e)}")
        return False
    finally:
        conn.close()


def mark_payment(athlete_id, month_year):
    conn = connect_db()
    if not conn:
        return False
    try:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO Payments (athlete_id, month_year, paid) 
            VALUES (?, ?, ?)
        """, (athlete_id, month_year, True))
        conn.commit()
        return True
    except pyodbc.IntegrityError:
        messagebox.showwarning("Ошибка", "Оплата за этот месяц уже зарегистрирована")
        return False
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка отметки оплаты: {str(e)}")
        return False
    finally:
        conn.close()


def get_payments_by_month(month_year, group_id=None):
    conn = connect_db()
    if not conn:
        return []
    try:
        cursor = conn.cursor()
        if group_id:
            cursor.execute("""
                SELECT A.athlete_id, A.name, 
                       CASE WHEN P.paid = True THEN 'Да' ELSE 'Нет' END as paid
                FROM Athletes A
                LEFT JOIN Payments P ON A.athlete_id = P.athlete_id AND P.month_year = ?
                WHERE A.current_group_id = ?
                ORDER BY A.name
            """, (month_year, group_id))
        else:
            cursor.execute("""
                SELECT A.athlete_id, A.name, 
                       CASE WHEN P.paid = True THEN 'Да' ELSE 'Нет' END as paid
                FROM Athletes A
                LEFT JOIN Payments P ON A.athlete_id = P.athlete_id AND P.month_year = ?
                ORDER BY A.name
            """, (month_year,))
        return cursor.fetchall()
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка загрузки платежей: {str(e)}")
        return []
    finally:
        conn.close()


def get_payment_stats(year):
    conn = connect_db()
    if not conn:
        return []
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT 
                SUBSTRING(month_year, 6, 2) as month,
                COUNT(*) as payment_count
            FROM Payments
            WHERE SUBSTRING(month_year, 1, 4) = ? AND paid = True
            GROUP BY SUBSTRING(month_year, 6, 2)
            ORDER BY month
        """, (year,))
        return cursor.fetchall()
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка загрузки статистики: {str(e)}")
        return []
    finally:
        conn.close()


def get_all_athletes_for_payment(group_id):
    conn = connect_db()
    if not conn:
        return []
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT athlete_id, name 
            FROM Athletes 
            WHERE current_group_id = ?
            ORDER BY name
        """, (group_id,))
        return cursor.fetchall()
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка загрузки спортсменов: {str(e)}")
        return []
    finally:
        conn.close()


# --- 4. GUI приложение ---
class SportClubApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Управление спортивным клубом")
        self.root.geometry("1200x800")

        self.current_group_id = None
        self.current_group_name = ""

        self.create_widgets()
        self.load_groups()
        self.update_stats()

    def create_widgets(self):
        # Главный фрейм
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Фрейм групп
        group_frame = tk.LabelFrame(main_frame, text="Группы", padx=5, pady=5)
        group_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5)

        self.group_listbox = tk.Listbox(group_frame, width=30, height=20)
        self.group_listbox.pack(fill=tk.BOTH, expand=True)
        self.group_listbox.bind('<<ListboxSelect>>', self.on_group_select)

        btn_frame = tk.Frame(group_frame)
        btn_frame.pack(fill=tk.X, pady=5)

        tk.Button(btn_frame, text="Добавить", command=self.add_group_dialog).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="Изменить", command=self.edit_group_dialog).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="Удалить", command=self.delete_group).pack(side=tk.LEFT, padx=2)

        # Notebook для остальных вкладок
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5)

        # Вкладка спортсменов
        self.create_athletes_tab()

        # Вкладка оплат
        self.create_payments_tab()

        # Вкладка статистики
        self.create_stats_tab()

        # Статус бар
        self.status_bar = tk.Label(self.root, text="Готово", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(fill=tk.X)

    def create_athletes_tab(self):
        tab = tk.Frame(self.notebook)
        self.notebook.add(tab, text="Спортсмены")

        # Поиск
        search_frame = tk.Frame(tab)
        search_frame.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(search_frame, text="Поиск:").pack(side=tk.LEFT)
        self.search_entry = tk.Entry(search_frame)
        self.search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.search_entry.bind('<KeyRelease>', lambda e: self.search_athletes())

        # Таблица спортсменов
        columns = ("ID", "ФИО", "Дата рождения", "Телефон")
        self.athletes_tree = ttk.Treeview(tab, columns=columns, show='headings', height=15)

        for col in columns:
            self.athletes_tree.heading(col, text=col)
            self.athletes_tree.column(col, width=50 if col == "ID" else 150)

        self.athletes_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Кнопки управления
        btn_frame = tk.Frame(tab)
        btn_frame.pack(fill=tk.X, pady=5)

        tk.Button(btn_frame, text="Добавить", command=self.add_athlete_dialog).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="Изменить", command=self.edit_athlete_dialog).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="Перевести", command=self.move_athlete_dialog).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="Удалить", command=self.delete_athlete).pack(side=tk.LEFT, padx=2)

    def create_payments_tab(self):
        tab = tk.Frame(self.notebook)
        self.notebook.add(tab, text="Оплаты")

        # Выбор месяца
        month_frame = tk.Frame(tab)
        month_frame.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(month_frame, text="Месяц:").pack(side=tk.LEFT)
        self.month_combo = ttk.Combobox(month_frame, state="readonly")
        self.month_combo.pack(side=tk.LEFT, padx=5)
        self.update_month_combo()

        # Выбор спортсмена
        tk.Label(month_frame, text="Спортсмен:").pack(side=tk.LEFT, padx=(10, 0))
        self.athlete_combo = ttk.Combobox(month_frame, state="readonly")
        self.athlete_combo.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # Таблица оплат
        columns = ("ФИО", "Оплачено")
        self.payments_tree = ttk.Treeview(tab, columns=columns, show='headings', height=15)

        for col in columns:
            self.payments_tree.heading(col, text=col)
            self.payments_tree.column(col, width=150)

        self.payments_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Кнопки управления
        btn_frame = tk.Frame(tab)
        btn_frame.pack(fill=tk.X, pady=5)

        tk.Button(btn_frame, text="Отметить оплату", command=self.mark_payment).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="Обновить", command=self.update_payments).pack(side=tk.LEFT, padx=2)

    def create_stats_tab(self):
        tab = tk.Frame(self.notebook)
        self.notebook.add(tab, text="Статистика")

        # Выбор года
        year_frame = tk.Frame(tab)
        year_frame.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(year_frame, text="Год:").pack(side=tk.LEFT)
        self.year_combo = ttk.Combobox(year_frame, state="readonly")
        self.year_combo.pack(side=tk.LEFT, padx=5)
        self.update_year_combo()

        # Таблица статистики
        columns = ("Месяц", "Кол-во оплат")
        self.stats_tree = ttk.Treeview(tab, columns=columns, show='headings', height=15)

        for col in columns:
            self.stats_tree.heading(col, text=col)
            self.stats_tree.column(col, width=150)

        self.stats_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Кнопки управления
        btn_frame = tk.Frame(tab)
        btn_frame.pack(fill=tk.X, pady=5)

        tk.Button(btn_frame, text="Обновить", command=self.update_stats).pack(side=tk.LEFT, padx=2)

    def update_month_combo(self):
        months = []
        current = datetime.now()
        for i in range(-6, 12):  # Прошлые 6 и будущие 12 месяцев
            date = current + timedelta(days=30 * i)
            months.append(date.strftime("%Y-%m"))
        self.month_combo['values'] = months
        self.month_combo.set(current.strftime("%Y-%m"))

    def update_year_combo(self):
        current_year = datetime.now().year
        years = [str(year) for year in range(current_year - 2, current_year + 3)]
        self.year_combo['values'] = years
        self.year_combo.set(str(current_year))

    def load_groups(self):
        self.groups = load_groups()
        self.group_listbox.delete(0, tk.END)
        for group in self.groups:
            self.group_listbox.insert(tk.END, group.group_name)

    def on_group_select(self, event):
        selection = self.group_listbox.curselection()
        if not selection:
            return

        self.current_group_id = self.groups[selection[0]].group_id
        self.current_group_name = self.groups[selection[0]].group_name
        self.update_athletes()
        self.update_payments()
        self.update_athlete_combo()
        self.update_status(f"Выбрана группа: {self.current_group_name}")

    def update_athletes(self):
        self.athletes_tree.delete(*self.athletes_tree.get_children())
        athletes = load_athletes(self.current_group_id)
        for athlete in athletes:
            birth_date = athlete.birth_date.strftime("%d.%m.%Y") if athlete.birth_date else ""
            self.athletes_tree.insert("", tk.END,
                                      values=(athlete.athlete_id, athlete.name, birth_date, athlete.phone))

    def update_athlete_combo(self):
        athletes = get_all_athletes_for_payment(self.current_group_id)
        self.athlete_combo['values'] = [a.name for a in athletes]
        if athletes:
            self.athlete_combo.current(0)

    def search_athletes(self):
        query = self.search_entry.get().lower()
        if not query:
            self.update_athletes()
            return

        self.athletes_tree.delete(*self.athletes_tree.get_children())
        athletes = load_athletes(self.current_group_id)
        for athlete in athletes:
            if query in athlete.name.lower():
                birth_date = athlete.birth_date.strftime("%d.%m.%Y") if athlete.birth_date else ""
                self.athletes_tree.insert("", tk.END,
                                          values=(athlete.athlete_id, athlete.name, birth_date, athlete.phone))

    def update_payments(self):
        if not self.current_group_id:
            return

        month = self.month_combo.get()
        self.payments_tree.delete(*self.payments_tree.get_children())

        payments = get_payments_by_month(month, self.current_group_id)
        for payment in payments:
            self.payments_tree.insert("", tk.END,
                                      values=(payment.name, payment.paid))

    def update_stats(self):
        year = self.year_combo.get()
        self.stats_tree.delete(*self.stats_tree.get_children())

        stats = get_payment_stats(year)
        month_names = [
            "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
            "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
        ]

        for stat in stats:
            month_num = int(stat.month)
            if 1 <= month_num <= 12:
                month_name = month_names[month_num - 1]
                self.stats_tree.insert("", tk.END,
                                       values=(month_name, stat.payment_count))

    def update_status(self, message):
        self.status_bar.config(text=message)

    # Диалоги и обработчики действий
    def add_group_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить группу")

        tk.Label(dialog, text="Название:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        name_entry = tk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(dialog, text="Описание:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
        desc_entry = tk.Text(dialog, width=30, height=5)
        desc_entry.grid(row=1, column=1, padx=5, pady=5)

        def save():
            name = name_entry.get().strip()
            desc = desc_entry.get("1.0", tk.END).strip()

            if not name:
                messagebox.showwarning("Ошибка", "Введите название группы")
                return

            if add_group(name, desc):
                self.load_groups()
                dialog.destroy()

        tk.Button(dialog, text="Сохранить", command=save).grid(row=2, column=1, pady=10)

    def edit_group_dialog(self):
        selection = self.group_listbox.curselection()
        if not selection:
            messagebox.showwarning("Ошибка", "Выберите группу для редактирования")
            return

        group = self.groups[selection[0]]

        dialog = tk.Toplevel(self.root)
        dialog.title("Изменить группу")

        tk.Label(dialog, text="Название:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        name_entry = tk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=5, pady=5)
        name_entry.insert(0, group.group_name)

        tk.Label(dialog, text="Описание:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
        desc_entry = tk.Text(dialog, width=30, height=5)
        desc_entry.grid(row=1, column=1, padx=5, pady=5)
        desc_entry.insert("1.0", group.description)

        def save():
            name = name_entry.get().strip()
            desc = desc_entry.get("1.0", tk.END).strip()

            if not name:
                messagebox.showwarning("Ошибка", "Введите название группы")
                return

            if update_group(group.group_id, name, desc):
                self.load_groups()
                dialog.destroy()

        tk.Button(dialog, text="Сохранить", command=save).grid(row=2, column=1, pady=10)

    def delete_group(self):
        selection = self.group_listbox.curselection()
        if not selection:
            messagebox.showwarning("Ошибка", "Выберите группу для удаления")
            return

        group = self.groups[selection[0]]

        if not messagebox.askyesno("Подтверждение",
                                   f"Удалить группу '{group.group_name}'? Все спортсмены из этой группы также будут удалены!"):
            return

        if delete_group(group.group_id):
            self.load_groups()
            self.current_group_id = None
            self.athletes_tree.delete(*self.athletes_tree.get_children())
            self.payments_tree.delete(*self.payments_tree.get_children())
            self.update_status(f"Группа '{group.group_name}' удалена")

    def add_athlete_dialog(self):
        if not self.current_group_id:
            messagebox.showwarning("Ошибка", "Выберите группу")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить спортсмена")

        tk.Label(dialog, text="ФИО:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        name_entry = tk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(dialog, text="Дата рождения:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
        birth_entry = DateEntry(dialog, width=12, date_pattern='dd.MM.yyyy')
        birth_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)

        tk.Label(dialog, text="Телефон:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.E)
        phone_entry = tk.Entry(dialog, width=30)
        phone_entry.grid(row=2, column=1, padx=5, pady=5)

        def save():
            name = name_entry.get().strip()
            birth = birth_entry.get_date().strftime("%Y-%m-%d")
            phone = phone_entry.get().strip()

            if not name:
                messagebox.showwarning("Ошибка", "Введите ФИО спортсмена")
                return

            if add_athlete(name, birth, phone, self.current_group_id):
                self.update_athletes()
                self.update_athlete_combo()
                dialog.destroy()

        tk.Button(dialog, text="Сохранить", command=save).grid(row=3, column=1, pady=10)

    def edit_athlete_dialog(self):
        selection = self.athletes_tree.selection()
        if not selection:
            messagebox.showwarning("Ошибка", "Выберите спортсмена для редактирования")
            return

        athlete_id = self.athletes_tree.item(selection)['values'][0]
        name = self.athletes_tree.item(selection)['values'][1]
        birth = self.athletes_tree.item(selection)['values'][2]
        phone = self.athletes_tree.item(selection)['values'][3]

        dialog = tk.Toplevel(self.root)
        dialog.title("Изменить спортсмена")

        tk.Label(dialog, text="ФИО:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        name_entry = tk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=5, pady=5)
        name_entry.insert(0, name)

        tk.Label(dialog, text="Дата рождения:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
        birth_entry = DateEntry(dialog, width=12, date_pattern='dd.MM.yyyy')
        birth_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        if birth:
            try:
                birth_entry.set_date(datetime.strptime(birth, "%d.%m.%Y"))
            except:
                pass

        tk.Label(dialog, text="Телефон:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.E)
        phone_entry = tk.Entry(dialog, width=30)
        phone_entry.grid(row=2, column=1, padx=5, pady=5)
        phone_entry.insert(0, phone)

        def save():
            new_name = name_entry.get().strip()
            new_birth = birth_entry.get_date().strftime("%Y-%m-%d")
            new_phone = phone_entry.get().strip()

            if not new_name:
                messagebox.showwarning("Ошибка", "Введите ФИО спортсмена")
                return

            if update_athlete(athlete_id, new_name, new_birth, new_phone, self.current_group_id):
                self.update_athletes()
                self.update_athlete_combo()
                dialog.destroy()

        tk.Button(dialog, text="Сохранить", command=save).grid(row=3, column=1, pady=10)

    def move_athlete_dialog(self):
        selection = self.athletes_tree.selection()
        if not selection:
            messagebox.showwarning("Ошибка", "Выберите спортсмена для перевода")
            return

        athlete_id = self.athletes_tree.item(selection)['values'][0]
        athlete_name = self.athletes_tree.item(selection)['values'][1]

        dialog = tk.Toplevel(self.root)
        dialog.title("Перевести спортсмена")

        tk.Label(dialog, text=f"Перевести {athlete_name} в:").pack(padx=10, pady=5)

        groups_listbox = tk.Listbox(dialog)
        groups_listbox.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        # Заполняем список всех групп, кроме текущей
        for group in self.groups:
            if group.group_id != self.current_group_id:
                groups_listbox.insert(tk.END, group.group_name)

        def move():
            selection = groups_listbox.curselection()
            if not selection:
                messagebox.showwarning("Ошибка", "Выберите группу для перевода")
                return

            # Находим ID выбранной группы
            selected_group = None
            idx = 0
            for group in self.groups:
                if group.group_id != self.current_group_id:
                    if idx == selection[0]:
                        selected_group = group
                        break
                    idx += 1

            if not selected_group:
                return

            if update_athlete(athlete_id, athlete_name,
                              self.athletes_tree.item(selection)['values'][2],  # birth_date
                              self.athletes_tree.item(selection)['values'][3],  # phone
                              selected_group.group_id):
                self.update_athletes()
                self.update_athlete_combo()
                self.update_status(f"Спортсмен {athlete_name} переведен в группу {selected_group.group_name}")
                dialog.destroy()

        tk.Button(dialog, text="Перевести", command=move).pack(pady=10)

    def delete_athlete(self):
        selection = self.athletes_tree.selection()
        if not selection:
            messagebox.showwarning("Ошибка", "Выберите спортсмена для удаления")
            return

        athlete_id = self.athletes_tree.item(selection)['values'][0]
        name = self.athletes_tree.item(selection)['values'][1]

        if not messagebox.askyesno("Подтверждение", f"Удалить спортсмена '{name}'?"):
            return

        if delete_athlete(athlete_id):
            self.update_athletes()
            self.update_athlete_combo()
            self.update_payments()
            self.update_status(f"Спортсмен '{name}' удален")

    def mark_payment(self):
        if not self.current_group_id:
            messagebox.showwarning("Ошибка", "Выберите группу")
            return

        month = self.month_combo.get()
        athlete_name = self.athlete_combo.get()

        if not athlete_name:
            messagebox.showwarning("Ошибка", "Выберите спортсмена")
            return

        # Находим ID спортсмена
        conn = connect_db()
        if not conn:
            return

        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT athlete_id 
                FROM Athletes 
                WHERE name = ? AND current_group_id = ?
            """, (athlete_name, self.current_group_id))
            result = cursor.fetchone()
            if not result:
                messagebox.showerror("Ошибка", "Спортсмен не найден")
                return

            athlete_id = result[0]

            if mark_payment(athlete_id, month):
                self.update_payments()
                self.update_stats()
                self.update_status(f"Оплата для '{athlete_name}' за {month} отмечена")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
        finally:
            conn.close()


# --- Запуск приложения ---
if __name__ == "__main__":
    root = tk.Tk()
    app = SportClubApp(root)
    root.mainloop()

