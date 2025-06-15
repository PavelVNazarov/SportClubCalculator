import os
import pyodbc
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime, timedelta
from tkcalendar import DateEntry
import pythoncom
import win32com.client
import shutil

# --- 1. Создание .accdb ---
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
        try:
            access_app.CloseCurrentDatabase()
        except:
            pass
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
            paid YESNO
        );
        """
        db.Execute(sql_payments)

        access_app.Quit()
        del access_app
        pythoncom.CoUninitialize()
        print(f"База данных создана: {db_path}")
        return True
    except Exception as e:
        print(f"Ошибка при создании БД: {e}")
        try:
            access_app.Quit()
        except:
            pass
        pythoncom.CoUninitialize()
        return False


# --- 2. Резервное копирование ---
def backup_database(db_path):
    try:
        backup_dir = os.path.join(os.path.dirname(db_path), "backups")
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(backup_dir, f"backup_{timestamp}.accdb")
        shutil.copy2(db_path, backup_path)
        print(f"Создана резервная копия: {backup_path}")
        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось создать резервную копию: {str(e)}")
        return False


# --- 3. Подключение к БД ---
def connect_db():
    db_path = os.path.join(os.path.dirname(__file__), "database.accdb")
    if not os.path.exists(db_path):
        success = create_access_database(db_path)
        if not success:
            messagebox.showerror("Ошибка", "Не удалось создать базу данных.")
            return None
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={db_path};'
    )
    try:
        return pyodbc.connect(conn_str)
    except Exception as e:
        messagebox.showerror("Ошибка подключения", str(e))
        return None


# --- 4. Функции работы с данными ---
def load_groups():
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("SELECT group_id, group_name, description FROM Groups")
    groups = cursor.fetchall()
    conn.close()
    return groups


def load_athletes(group_id=None):
    conn = connect_db()
    cursor = conn.cursor()
    if group_id:
        cursor.execute("SELECT athlete_id, name, birth_date, phone FROM Athletes WHERE current_group_id = ?", (group_id,))
    else:
        cursor.execute("SELECT athlete_id, name, birth_date, phone FROM Athletes")
    athletes = cursor.fetchall()
    conn.close()
    return athletes


def add_group(name, desc):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("INSERT INTO Groups (group_name, description) VALUES (?, ?)", (name, desc))
    conn.commit()
    conn.close()


def delete_group(gid):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM Groups WHERE group_id = ?", (gid,))
    conn.commit()
    conn.close()


def add_athlete(name, birth, phone, gid):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("INSERT INTO Athletes (name, birth_date, phone, current_group_id) VALUES (?, ?, ?, ?)",
                   (name, birth, phone, gid))
    conn.commit()
    conn.close()


def move_athlete(aid, new_gid):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("UPDATE Athletes SET current_group_id = ? WHERE athlete_id = ?", (new_gid, aid))
    conn.commit()
    conn.close()


def delete_athlete(aid):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM Athletes WHERE athlete_id = ?", (aid,))
    conn.commit()
    conn.close()


def mark_payment(aid, month):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Payments WHERE athlete_id = ? AND month_year = ?", (aid, month))
    if cursor.fetchone():
        conn.close()
        return False  # Оплата уже есть
    cursor.execute("INSERT INTO Payments (athlete_id, month_year, paid) VALUES (?, ?, ?)", (aid, month, True))
    conn.commit()
    conn.close()
    return True


def get_all_payments_by_months(group_id=None):
    conn = connect_db()
    cursor = conn.cursor()
    query = """
        SELECT P.month_year, A.name, G.group_name, P.paid 
        FROM ((Payments AS P 
        INNER JOIN Athletes AS A ON P.athlete_id = A.athlete_id)
        INNER JOIN Groups AS G ON A.current_group_id = G.group_id)
    """
    params = []
    if group_id:
        query += " WHERE A.current_group_id = ?"
        params.append(group_id)
    query += " ORDER BY P.month_year DESC"
    cursor.execute(query, params)
    res = cursor.fetchall()
    conn.close()
    return res


def get_payments_by_month(month, group_id):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT A.name, P.paid 
        FROM (Payments AS P INNER JOIN Athletes AS A ON P.athlete_id = A.athlete_id)
        WHERE P.month_year = ? AND A.current_group_id = ?
    """, (month, group_id))
    res = cursor.fetchall()
    conn.close()
    return res


def get_unpaid_athletes(month, group_id):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT A.name FROM Athletes A
        WHERE NOT EXISTS (
            SELECT 1 FROM Payments P
            WHERE P.athlete_id = A.athlete_id AND P.month_year = ?
        ) AND A.current_group_id = ?
    """, (month, group_id))
    res = cursor.fetchall()
    conn.close()
    return [r.name for r in res]


def get_payment_stats(year):
    conn = connect_db()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT 
                SUBSTR(month_year, 7, 2) AS month,
                COUNT(*) AS count
            FROM Payments
            WHERE SUBSTR(month_year, 1, 4) = ? AND paid = True
            GROUP BY SUBSTR(month_year, 7, 2)
            ORDER BY month
        """, (year,))
        result = cursor.fetchall()
        conn.close()
        return result
    except Exception as e:
        conn.close()
        print(f"[Ошибка] При загрузке статистики: {e}")
        return []


def export_to_excel(data, columns, filename):
    df = pd.DataFrame(data, columns=columns)
    df.to_excel(filename, index=False)
    messagebox.showinfo("Экспорт", f"Сохранено в {filename}")


def export_all_to_excel(data_dict, filename):
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        for sheet_name, (data, columns) in data_dict.items():
            if data:
                df = pd.DataFrame(data, columns=columns)
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    messagebox.showinfo("Экспорт", f"Данные сохранены в {filename}")


# --- 5. Интерфейс Tkinter ---
class SportClubApp:
    def __init__(self, root):
        self.root = root
        self.root.title("СпортКлуб — Управление")
        self.root.geometry("1000x600")

        self.current_group_id = None
        self.groups_data = []

        # Список групп
        self.group_listbox = tk.Listbox(root)
        self.group_listbox.pack(side="left", fill="y", padx=5, pady=5)
        self.group_listbox.bind("<<ListboxSelect>>", self.on_group_select)

        # Основной интерфейс
        self.frame = tk.Frame(root)
        self.frame.pack(side="right", fill="both", expand=True)

        self.create_widgets()
        self.load_groups()

    def create_widgets(self):
        self.notebook = ttk.Notebook(self.frame)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)

        # Вкладка "Группы"
        tab_groups = tk.Frame(self.notebook)
        self.notebook.add(tab_groups, text="Группы")
        tk.Button(tab_groups, text="Добавить группу", command=self.add_new_group).pack(pady=5)
        tk.Button(tab_groups, text="Удалить группу", command=self.delete_selected_group).pack(pady=5)

        # Вкладка "Спортсмены"
        tab_athletes = tk.Frame(self.notebook)
        self.notebook.add(tab_athletes, text="Спортсмены")

        self.group_description_label = tk.Label(tab_athletes, text="", wraplength=400, justify="left")
        self.group_description_label.pack(pady=5)

        self.tree_athletes = ttk.Treeview(tab_athletes, columns=("ID", "ФИО", "Дата рождения", "Телефон"), show='headings')
        for col in ("ID", "ФИО", "Дата рождения", "Телефон"):
            self.tree_athletes.heading(col, text=col)
            self.tree_athletes.column(col, width=80 if col == "ID" else 150)
        self.tree_athletes.pack(fill="both", expand=True, padx=5, pady=5)

        search_frame = tk.Frame(tab_athletes)
        search_frame.pack(pady=5)
        tk.Label(search_frame, text="Поиск по ФИО:").pack(side=tk.LEFT)
        self.search_entry = tk.Entry(search_frame, width=30)
        self.search_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(search_frame, text="Найти", command=self.search_athlete).pack(side=tk.LEFT)
        tk.Button(tab_athletes, text="Добавить спортсмена", command=self.add_new_athlete).pack(pady=5)
        tk.Button(tab_athletes, text="Перевести в группу", command=self.move_selected_athlete).pack(pady=5)
        tk.Button(tab_athletes, text="Удалить спортсмена", command=self.delete_selected_athlete).pack(pady=5)

        # Вкладка "Оплата"
        tab_payments = tk.Frame(self.notebook)
        self.notebook.add(tab_payments, text="Оплата")

        frame = tk.Frame(tab_payments)
        frame.pack(pady=10)
        tk.Label(frame, text="Месяц (YYYY-MM):").grid(row=0, column=0)
        self.month_selector = ttk.Combobox(frame, values=fill_month_selector(), state="readonly")
        self.month_selector.grid(row=0, column=1)
        self.month_selector.set(datetime.now().strftime("%Y-%m"))
        tk.Button(frame, text="Отметить как оплачено", command=self.mark_selected_payment).grid(row=0, column=2, padx=5)

        self.athlete_selector = ttk.Combobox(tab_payments, state="readonly")
        self.athlete_selector.pack(padx=10, pady=5)

        self.tree_payments = ttk.Treeview(tab_payments, columns=("Имя", "Оплачено"), show='headings')
        self.tree_payments.heading("Имя", text="Имя")
        self.tree_payments.heading("Оплачено", text="Оплачено")
        self.tree_payments.pack(fill="both", expand=True, padx=5, pady=5)

        tk.Button(tab_payments, text="Показать оплату за месяц", command=self.show_payments).pack(pady=5)
        tk.Button(tab_payments, text="Экспорт в Excel", command=self.export_payments).pack(pady=5)

        # Вкладка "Все оплаты"
        tab_all_payments = tk.Frame(self.notebook)
        self.notebook.add(tab_all_payments, text="Все оплаты")

        filter_frame = tk.Frame(tab_all_payments)
        filter_frame.pack(pady=10)

        tk.Label(filter_frame, text="Группа:").grid(row=0, column=0)
        self.all_payments_group = ttk.Combobox(filter_frame, state="readonly")
        self.all_payments_group.grid(row=0, column=1)

        tk.Label(filter_frame, text="С:").grid(row=0, column=2, padx=(10, 0))
        self.date_from = DateEntry(filter_frame, date_pattern='yyyy-mm-dd', width=10)
        self.date_from.grid(row=0, column=3)

        tk.Label(filter_frame, text="По:").grid(row=0, column=4, padx=(10, 0))
        self.date_to = DateEntry(filter_frame, date_pattern='yyyy-mm-dd', width=10)
        self.date_to.grid(row=0, column=5)

        tk.Button(filter_frame, text="Обновить", command=self.load_all_payments).grid(row=0, column=6, padx=10)

        self.tree_all_payments = ttk.Treeview(tab_all_payments,
                                             columns=("Месяц", "Имя", "Группа", "Оплачено"),
                                             show='headings')
        for col in ("Месяц", "Имя", "Группа", "Оплачено"):
            self.tree_all_payments.heading(col, text=col)
        self.tree_all_payments.pack(fill="both", expand=True, padx=5, pady=5)

        tk.Button(tab_all_payments, text="Экспорт всех данных в Excel",
                  command=self.export_all_data).pack(pady=5)

        # Кнопка резервного копирования
        tk.Button(tab_all_payments, text="Сделать резервную копию", command=self.backup_current_db).pack(pady=5)

        # Привязка переключения вкладок
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

    def on_group_select(self, event):
        selected = self.group_listbox.curselection()
        if not selected:
            return
        group = self.groups_data[selected[0]]
        self.current_group_id = group.group_id
        self.group_description_label.config(text=f"{group.description}")
        self.load_athletes_in_group(group.group_id)
        self.load_athlete_names_for_payment()

    def on_tab_change(self, event):
        current_tab = self.notebook.tab(self.notebook.select(), "text")
        if current_tab == "Оплата" and self.current_group_id:
            self.load_athlete_names_for_payment()
        elif current_tab == "Все оплаты":
            self.load_groups_for_filter()

    def load_groups(self):
        self.group_listbox.delete(0, tk.END)
        self.groups_data = load_groups()
        for g in self.groups_data:
            self.group_listbox.insert(tk.END, g.group_name)

    def load_athletes_in_group(self, group_id):
        self.tree_athletes.delete(*self.tree_athletes.get_children())
        athletes = load_athletes(group_id)
        for a in athletes:
            birth = a.birth_date.strftime("%d.%m.%Y") if a.birth_date else ""
            self.tree_athletes.insert("", tk.END, values=(a.athlete_id, a.name, birth, a.phone))

    def load_athlete_names_for_payment(self):
        athletes = load_athletes(self.current_group_id)
        names = [a.name for a in athletes]
        self.athlete_selector['values'] = names
        if names:
            self.athlete_selector.current(0)

    def add_new_group(self):
        name = simple_input("Введите название группы:")
        if not name:
            return
        desc = simple_input("Введите описание группы:")
        add_group(name, desc)
        self.load_groups()

    def delete_selected_group(self):
        if self.current_group_id is None:
            messagebox.showwarning("Ошибка", "Выберите группу!")
            return
        answer = messagebox.askyesno("Подтверждение", "Вы действительно хотите удалить группу?")
        if not answer:
            return
        group = self.groups_data[self.group_listbox.curselection()[0]]
        delete_group(group.group_id)
        self.load_groups()

    def add_new_athlete(self):
        if self.current_group_id is None:
            messagebox.showwarning("Ошибка", "Выберите группу!")
            return
        data = ask_name_birth_phone()
        if not data or not data['name'] or not data['phone']:
            return
        name = data['name']
        birth = data['birth']
        phone = data['phone']
        add_athlete(name, birth, phone, self.current_group_id)
        self.load_athletes_in_group(self.current_group_id)

    def move_selected_athlete(self):
        selected = self.tree_athletes.selection()
        if not selected:
            messagebox.showwarning("Ошибка", "Выберите спортсмена!")
            return
        athlete_id = self.tree_athletes.item(selected)['values'][0]
        new_group = select_group_dialog(self.groups_data)
        if new_group:
            move_athlete(athlete_id, new_group)
            self.load_athletes_in_group(new_group)
            self.load_athletes_in_group(self.current_group_id)

    def delete_selected_athlete(self):
        selected = self.tree_athletes.selection()
        if not selected:
            messagebox.showwarning("Ошибка", "Выберите спортсмена!")
            return
        athlete_id = self.tree_athletes.item(selected)['values'][0]
        name = self.tree_athletes.item(selected)['values'][1]
        answer = messagebox.askyesno("Подтверждение", f"Удалить спортсмена {name}?")
        if not answer:
            return
        delete_athlete(athlete_id)
        self.load_athletes_in_group(self.current_group_id)

    def mark_selected_payment(self):
        selected_name = self.athlete_selector.get()
        if not selected_name:
            messagebox.showwarning("Ошибка", "Выберите спортсмена!")
            return
        conn = connect_db()
        cursor = conn.cursor()
        cursor.execute("SELECT athlete_id FROM Athletes WHERE name = ?", (selected_name,))
        row = cursor.fetchone()
        if not row:
            messagebox.showerror("Ошибка", "Спортсмен не найден!")
            return
        athlete_id = row[0]
        month = self.month_selector.get()
        cursor.execute("SELECT * FROM Payments WHERE athlete_id = ? AND month_year = ?", (athlete_id, month))
        if cursor.fetchone():
            messagebox.showinfo("Информация", "Оплата уже отмечена.")
            return
        cursor.execute("INSERT INTO Payments (athlete_id, month_year, paid) VALUES (?, ?, ?)", (athlete_id, month, True))
        conn.commit()
        conn.close()
        self.show_payments()

    def show_payments(self):
        month = self.month_selector.get()
        payments = get_payments_by_month(month, self.current_group_id)
        unpaid = get_unpaid_athletes(month, self.current_group_id)
        self.tree_payments.delete(*self.tree_payments.get_children())
        for p in payments:
            self.tree_payments.insert("", tk.END, values=(p.name, "Да" if p.paid else "Нет"))
        for name in unpaid:
            self.tree_payments.insert("", tk.END, values=(name, "Нет"))

    def export_payments(self):
        month = self.month_selector.get()
        payments = get_payments_by_month(month, self.current_group_id)
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if filename:
            export_to_excel([(p.name, "Да" if p.paid else "Нет") for p in payments],
                            ["Имя", "Оплачено"], filename)

    def load_all_payments(self):
        group_name = self.all_payments_group.get()
        group_id = None
        if group_name != 'Все':
            group = next((g for g in self.groups_data if g.group_name == group_name), None)
            if group:
                group_id = group.group_id

        date_from = self.date_from.get_date().strftime("%Y-%m")
        date_to = self.date_to.get_date().strftime("%Y-%m")

        conn = connect_db()
        cursor = conn.cursor()
        query = """
            SELECT P.month_year, A.name, G.group_name, P.paid 
            FROM ((Payments AS P 
            INNER JOIN Athletes AS A ON P.athlete_id = A.athlete_id)
            INNER JOIN Groups AS G ON A.current_group_id = G.group_id)
            WHERE P.month_year BETWEEN ? AND ?
        """
        params = [date_from, date_to]

        if group_id:
            query += " AND A.current_group_id = ?"
            params.append(group_id)

        query += " ORDER BY P.month_year DESC"

        try:
            cursor.execute(query, params)
            res = cursor.fetchall()
            self.tree_all_payments.delete(*self.tree_all_payments.get_children())
            for row in res:
                paid_str = "Да" if row.paid else "Нет"
                self.tree_all_payments.insert("", tk.END, values=(row.month_year, row.name, row.group_name, paid_str))
        finally:
            conn.close()

    def load_groups_for_filter(self):
        groups = load_groups()
        names = [g.group_name for g in groups]
        self.all_payments_group['values'] = ['Все'] + names
        self.all_payments_group.set('Все')

    def export_all_data(self):
        if not self.current_group_id:
            messagebox.showwarning("Ошибка", "Выберите группу!")
            return

        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not filename:
            return

        try:
            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                # Спортсмены
                athletes = []
                for item in self.tree_athletes.get_children():
                    values = self.tree_athletes.item(item)['values']
                    if len(values) >= 4:
                        athletes.append({
                            "ID": values[0],
                            "ФИО": values[1],
                            "Дата рождения": values[2],
                            "Телефон": values[3]
                        })
                if athletes:
                    df = pd.DataFrame(athletes)
                    df.to_excel(writer, sheet_name="Спортсмены", index=False)

                # Оплаты
                payments = []
                month = self.month_selector.get()
                for item in self.tree_payments.get_children():
                    values = self.tree_payments.item(item)['values']
                    if len(values) >= 2:
                        payments.append({"ФИО": values[0], "Оплачено": values[1], "Месяц": month})
                if payments:
                    df = pd.DataFrame(payments)
                    df.to_excel(writer, sheet_name="Оплата", index=False)

                # Статистика
                stats = []
                for item in self.stats_tree.get_children():
                    values = self.stats_tree.item(item)['values']
                    if len(values) >= 2:
                        stats.append({"Месяц": values[0], "Количество оплат": values[1]})
                if stats:
                    df = pd.DataFrame(stats)
                    df.to_excel(writer, sheet_name="Статистика", index=False)

            messagebox.showinfo("Экспорт", f"Данные успешно сохранены в:\n{filename}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{str(e)}")

    def backup_current_db(self):
        db_path = os.path.join(os.path.dirname(__file__), "database.accdb")
        if os.path.exists(db_path):
            if backup_database(db_path):
                messagebox.showinfo("Резервная копия", "Резервная копия создана!")
        else:
            messagebox.showerror("Ошибка", "База данных не найдена.")

    def search_athlete(self):
        query = self.search_entry.get().strip().lower()
        if not query:
            self.load_athletes_in_group(self.current_group_id)
            return
        athletes = load_athletes(self.current_group_id)
        filtered = [a for a in athletes if query in a.name.lower()]
        self.tree_athletes.delete(*self.tree_athletes.get_children())
        for a in filtered:
            birth = a.birth_date.strftime("%d.%m.%Y") if a.birth_date else ""
            self.tree_athletes.insert("", tk.END, values=(a.athlete_id, a.name, birth, a.phone))


# --- Вспомогательные функции ---
def simple_input(prompt):
    result = []
    def on_ok():
        result.append(entry.get())
        dialog.destroy()
    dialog = tk.Toplevel()
    dialog.title("Ввод")
    tk.Label(dialog, text=prompt).pack(padx=10, pady=5)
    entry = tk.Entry(dialog, width=40)
    entry.pack(padx=10, pady=5)
    tk.Button(dialog, text="OK", command=on_ok).pack(pady=10)
    dialog.wait_window()
    return result[0] if result else ""


def ask_name_birth_phone():
    result = {}
    def on_ok():
        result.update({
            'name': entry_name.get(),
            'phone': entry_phone.get(),
            'birth': cal.get_date().strftime("%Y-%m-%d")
        })
        dialog.destroy()
    dialog = tk.Toplevel()
    dialog.title("Добавить спортсмена")
    tk.Label(dialog, text="ФИО:").pack(padx=10, pady=5)
    entry_name = tk.Entry(dialog, width=40)
    entry_name.pack(padx=10, pady=5)
    tk.Label(dialog, text="Дата рождения:").pack(padx=10, pady=5)
    cal = DateEntry(dialog, date_pattern='yyyy-mm-dd')
    cal.pack(padx=10, pady=5)
    tk.Label(dialog, text="Телефон:").pack(padx=10, pady=5)
    entry_phone = tk.Entry(dialog, width=30)
    entry_phone.pack(padx=10, pady=5)
    tk.Button(dialog, text="OK", command=on_ok).pack(pady=10)
    dialog.wait_window()
    return result


def select_group_dialog(groups):
    result = []
    dialog = tk.Toplevel()
    dialog.title("Выберите группу")
    listbox = tk.Listbox(dialog)
    listbox.pack(padx=10, pady=10)
    for g in groups:
        listbox.insert(tk.END, g.group_name)
    def on_ok():
        sel = listbox.curselection()
        if sel:
            result.append(groups[sel[0]].group_id)
        dialog.destroy()
    tk.Button(dialog, text="Выбрать", command=on_ok).pack()
    dialog.wait_window()
    return result[0] if result else None


def fill_month_selector():
    now = datetime.now()
    months = [(now - timedelta(days=30 * i)).strftime("%Y-%m") for i in range(24)]
    for i in range(1, 6):
        for m in range(1, 13):
            dt = now.replace(year=now.year + i, month=m, day=1)
            months.append(dt.strftime("%Y-%m"))
    return sorted(set(months), reverse=True)


# --- Запуск приложения ---
if __name__ == "__main__":
    db_path = os.path.join(os.path.dirname(__file__), "database.accdb")
    print("Путь к БД:", db_path)
    if not os.path.exists(db_path):
        print("Создаю новую базу данных...")
        success = create_access_database(db_path)
        if not success:
            messagebox.showerror("Ошибка", "Не удалось создать базу данных.")
            exit(1)

    root = tk.Tk()
    app = SportClubApp(root)
    root.mainloop()


