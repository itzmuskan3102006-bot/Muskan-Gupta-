"""
main.py - Student Grading System (Dark Themed Version)
Theme: VS Code Style Dark UI (Option A)
Single-file app with:
- SQLite persistent storage (students, subjects, marks, attendance)
- Tkinter GUI (dark theme)
- Add / Edit / Delete students and subjects
- Multi-subject marks entry
- Auto Total / Percentage / GPA / Grade
- Attendance marking & % calculation
- Search & Grade filter
- Export to CSV & Excel (optional, requires pandas + openpyxl)
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import sqlite3
from contextlib import closing
import datetime
import csv

# Optional Excel export
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except:
    PANDAS_AVAILABLE = False

DB_FILE = "students_dark.db"

# -----------------------------------------------------------
#  DARK THEME COLORS
# -----------------------------------------------------------

BG = "#1e1e1e"         # Main window background
FRAME_BG = "#252526"   # Frame background
BTN_BG = "#007acc"     # Blue buttons
BTN_HOVER = "#149eff"  # Hover color
TEXT_COLOR = "#ffffff" # White text
ENTRY_BG = "#3c3c3c"   # Input fields
ENTRY_FG = "#ffffff"
TREE_BG = "#1e1e1e"
TREE_FG = "#ffffff"
TREE_SELECTED = "#094771"
LISTBOX_BG = "#2d2d30"
LISTBOX_FG = "#ffffff"

# -----------------------------------------------------------
#  DATABASE FUNCTIONS
# -----------------------------------------------------------

def init_db():
    conn = sqlite3.connect(DB_FILE)
    with closing(conn):
        c = conn.cursor()

        c.execute('''
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id TEXT UNIQUE,
            name TEXT NOT NULL,
            klass TEXT
        )''')

        c.execute('''
        CREATE TABLE IF NOT EXISTS subjects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL
        )''')

        c.execute('''
        CREATE TABLE IF NOT EXISTS marks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_db_id INTEGER,
            subject_db_id INTEGER,
            marks REAL,
            FOREIGN KEY(student_db_id) REFERENCES students(id),
            FOREIGN KEY(subject_db_id) REFERENCES subjects(id)
        )''')

        c.execute('''
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_db_id INTEGER,
            date TEXT,
            present INTEGER,
            FOREIGN KEY(student_db_id) REFERENCES students(id)
        )''')

        conn.commit()


def get_conn():
    return sqlite3.connect(DB_FILE)


def add_student_to_db(student_id, name, klass):
    conn = get_conn()
    with closing(conn):
        c = conn.cursor()
        c.execute("INSERT INTO students (student_id, name, klass) VALUES (?, ?, ?)",
                  (student_id, name, klass))
        conn.commit()
        return c.lastrowid


def update_student_in_db(db_id, student_id, name, klass):
    conn = get_conn()
    with closing(conn):
        c = conn.cursor()
        c.execute("UPDATE students SET student_id=?, name=?, klass=? WHERE id=?",
                  (student_id, name, klass, db_id))
        conn.commit()


def delete_student_from_db(db_id):
    conn = get_conn()
    with closing(conn):
        c = conn.cursor()
        c.execute("DELETE FROM students WHERE id=?", (db_id,))
        c.execute("DELETE FROM marks WHERE student_db_id=?", (db_id,))
        c.execute("DELETE FROM attendance WHERE student_db_id=?", (db_id,))
        conn.commit()


def list_students_from_db():
    conn = get_conn()
    with closing(conn):
        c = conn.cursor()
        c.execute("SELECT id, student_id, name, klass FROM students ORDER BY name")
        return c.fetchall()


def find_students_by_name(name):
    conn = get_conn()
    with closing(conn):
        c = conn.cursor()
        c.execute("SELECT id, student_id, name, klass FROM students WHERE name LIKE ? ORDER BY name",
                  (f"%{name}%",))
        return c.fetchall()


def add_subject_to_db(name):
    conn = get_conn()
    with closing(conn):
        c = conn.cursor()
        c.execute("INSERT OR IGNORE INTO subjects (name) VALUES (?)", (name,))
        conn.commit()


def list_subjects_from_db():
    conn = get_conn()
    with closing(conn):
        c = conn.cursor()
        c.execute("SELECT id, name FROM subjects ORDER BY name")
        return c.fetchall()


def set_mark_in_db(student_db_id, subject_db_id, marks):
    conn = get_conn()
    with closing(conn):
        c = conn.cursor()
        c.execute("SELECT id FROM marks WHERE student_db_id=? AND subject_db_id=?",
                  (student_db_id, subject_db_id))
        row = c.fetchone()

        if row:
            c.execute("UPDATE marks SET marks=? WHERE id=?", (marks, row[0]))
        else:
            c.execute("INSERT INTO marks (student_db_id, subject_db_id, marks) VALUES (?, ?, ?)",
                      (student_db_id, subject_db_id, marks))

        conn.commit()


def get_marks_for_student_from_db(student_db_id):
    conn = get_conn()
    with closing(conn):
        c = conn.cursor()
        c.execute('''
        SELECT subjects.name, marks.marks
        FROM marks
        JOIN subjects ON subjects.id = marks.subject_db_id
        WHERE marks.student_db_id = ?
        ORDER BY subjects.name
        ''', (student_db_id,))
        return c.fetchall()


def delete_marks_for_subject(subject_db_id):
    conn = get_conn()
    with closing(conn):
        c = conn.cursor()
        c.execute("DELETE FROM marks WHERE subject_db_id=?", (subject_db_id,))
        conn.commit()


def add_attendance_to_db(student_db_id, date_str, present):
    conn = get_conn()
    with closing(conn):
        c = conn.cursor()
        c.execute("INSERT INTO attendance (student_db_id, date, present) VALUES (?, ?, ?)",
                  (student_db_id, date_str, 1 if present else 0))
        conn.commit()


def get_attendance_percent_from_db(student_db_id):
    conn = get_conn()
    with closing(conn):
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM attendance WHERE student_db_id=?", (student_db_id,))
        total = c.fetchone()[0]

        if total == 0:
            return 0.0

        c.execute("SELECT SUM(present) FROM attendance WHERE student_db_id=?", (student_db_id,))
        present = c.fetchone()[0] or 0

        return round((present / total) * 100, 2)


# -----------------------------------------------------------
#  GRADING LOGIC
# -----------------------------------------------------------

def calculate_total_percentage_gpa_grade(marks_list):
    if not marks_list:
        return 0, 0, 0, 0, "N/A"

    total = 0
    max_total = 100 * len(marks_list)

    for _, m in marks_list:
        try:
            total += float(m)
        except:
            total += 0

    percentage = round((total / max_total) * 100, 2) if max_total > 0 else 0.0

    if percentage >= 90:
        return total, max_total, percentage, 4.0, "A+"
    if percentage >= 80:
        return total, max_total, percentage, 3.7, "A"
    if percentage >= 70:
        return total, max_total, percentage, 3.0, "B"
    if percentage >= 60:
        return total, max_total, percentage, 2.0, "C"
    if percentage >= 50:
        return total, max_total, percentage, 1.0, "D"

    return total, max_total, percentage, 0.0, "F"


# -----------------------------------------------------------
#  GUI – DARK THEME IMPLEMENTATION
# -----------------------------------------------------------

class StudentGradingApp:
    def __init__(self, master):
        self.master = master
        master.title("Student Grading System — Dark Mode")
        master.geometry("1100x700")
        master.configure(bg=BG)
        master.protocol("WM_DELETE_WINDOW", self.on_close)

        # ---------------------------
        # Widget Style Overrides
        # ---------------------------

        style = ttk.Style()
        try:
            style.theme_use("default")
        except:
            pass

        # Treeview Dark Mode
        style.configure("Treeview",
                        background=TREE_BG,
                        foreground=TREE_FG,
                        fieldbackground=TREE_BG,
                        rowheight=28,
                        bordercolor=BG,
                        borderwidth=0,
                        font=("Segoe UI", 10))

        style.map("Treeview",
                  background=[("selected", TREE_SELECTED)])

        style.configure("Treeview.Heading",
                        background="#3c3c3c",
                        foreground="#ffffff",
                        font=("Segoe UI", 10, "bold"))

        # Combobox Dark Mode
        style.configure("TCombobox",
                        fieldbackground=ENTRY_BG,
                        background=ENTRY_BG,
                        foreground=ENTRY_FG)

        # ---------------------------
        # TOP SEARCH BAR
        # ---------------------------

        top_frame = tk.Frame(master, bg=FRAME_BG)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=8, pady=8)

        tk.Label(top_frame, text="Search Name:", fg=TEXT_COLOR, bg=FRAME_BG).pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(top_frame, textvariable=self.search_var,
                                     bg=ENTRY_BG, fg=ENTRY_FG, insertbackground="white")
        self.search_entry.pack(side=tk.LEFT, padx=6)
        self.search_entry.bind("<Return>", lambda e: self.refresh_student_list())

        self.create_button(top_frame, "Search", self.refresh_student_list).pack(side=tk.LEFT, padx=6)
        self.create_button(top_frame, "Clear Search", self.clear_search).pack(side=tk.LEFT, padx=6)

        tk.Label(top_frame, text="Filter Grade:", fg=TEXT_COLOR, bg=FRAME_BG).pack(side=tk.LEFT, padx=(20,0))

        self.filter_grade_var = tk.StringVar(value="All")
        grade_options = ["All", "A+", "A", "B", "C", "D", "F"]

        self.grade_box = ttk.Combobox(top_frame,
                                      textvariable=self.filter_grade_var,
                                      values=grade_options,
                                      state="readonly",
                                      width=6)
        self.grade_box.pack(side=tk.LEFT, padx=6)

        self.create_button(top_frame, "Apply Filter", self.refresh_student_list).pack(side=tk.LEFT, padx=6)

        # ---------------------------
        # LEFT FRAME (Student + Subjects + Export + Attendance)
        # ---------------------------

        left_frame = tk.Frame(master, bg=FRAME_BG)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        # Student Form
        tk.Label(left_frame, text="Student Details", fg=TEXT_COLOR,
                 bg=FRAME_BG, font=("Segoe UI", 12, "bold")).pack(anchor="w")

        form_frame = tk.Frame(left_frame, bg=FRAME_BG)
        form_frame.pack(fill=tk.X, pady=4)

        def create_label(text, row):
            tk.Label(form_frame, text=text, fg=TEXT_COLOR, bg=FRAME_BG)\
                .grid(row=row, column=0, sticky="w", pady=2)

        def create_entry(var, row):
            e = tk.Entry(form_frame, textvariable=var, bg=ENTRY_BG,
                         fg=ENTRY_FG, insertbackground="white")
            e.grid(row=row, column=1, sticky="we", pady=2)
            return e

        self.sid_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.class_var = tk.StringVar()

        create_label("Student ID:", 0)
        self.e_student_id = create_entry(self.sid_var, 0)

        create_label("Name:", 1)
        self.e_name = create_entry(self.name_var, 1)

        create_label("Class:", 2)
        self.e_class = create_entry(self.class_var, 2)

        button_bar = tk.Frame(left_frame, bg=FRAME_BG)
        button_bar.pack(fill=tk.X, pady=6)

        self.create_button(button_bar, "Add Student", self.add_student).pack(side=tk.LEFT, padx=3)
        self.create_button(button_bar, "Update Selected", self.update_selected_student).pack(side=tk.LEFT, padx=3)
        self.create_button(button_bar, "Delete Selected", self.delete_selected_student).pack(side=tk.LEFT, padx=3)
        self.create_button(button_bar, "Clear Form", self.clear_form).pack(side=tk.LEFT, padx=3)

        # ---------------------------
        # SUBJECTS PANEL
        # ---------------------------

        tk.Label(left_frame, text="Subjects", fg=TEXT_COLOR,
                 bg=FRAME_BG, font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(10,0))

        subj_box = tk.Frame(left_frame, bg=FRAME_BG)
        subj_box.pack(fill=tk.X)

        self.subject_name_var = tk.StringVar()
        self.e_subject_name = tk.Entry(subj_box, textvariable=self.subject_name_var,
                                       bg=ENTRY_BG, fg=ENTRY_FG, insertbackground="white")
        self.e_subject_name.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        self.create_button(subj_box, "Add", self.add_subject, w=6).pack(side=tk.LEFT)
        self.create_button(subj_box, "Edit", self.edit_subject, w=6).pack(side=tk.LEFT, padx=2)
        self.create_button(subj_box, "Delete", self.delete_subject, w=6).pack(side=tk.LEFT)

        tk.Label(left_frame, text="Subject List", fg=TEXT_COLOR,
                 bg=FRAME_BG).pack(anchor="w", pady=4)

        self.subj_listbox = tk.Listbox(left_frame, bg=LISTBOX_BG, fg=LISTBOX_FG,
                                       height=8, selectbackground=TREE_SELECTED)
        self.subj_listbox.pack(fill=tk.X)
        self.subj_listbox.bind("<<ListboxSelect>>", self.on_subject_select)

        # ---------------------------
        # ATTENDANCE PANEL
        # ---------------------------

        att_box = tk.LabelFrame(left_frame, text="Attendance", bg=FRAME_BG,
                                fg=TEXT_COLOR, labelanchor="n")
        att_box.pack(fill=tk.X, pady=8)

        self.create_button(att_box, "Mark Present Today", self.mark_present_today).pack(fill=tk.X, pady=3)
        self.create_button(att_box, "Show Attendance %", self.show_attendance_percent).pack(fill=tk.X, pady=3)

        # ---------------------------
        # EXPORT PANEL
        # ---------------------------

        export_box = tk.LabelFrame(left_frame, text="Export", bg=FRAME_BG,
                                   fg=TEXT_COLOR, labelanchor="n")
        export_box.pack(fill=tk.X, pady=8)

        self.create_button(export_box, "Export CSV", self.export_visible_csv).pack(fill=tk.X, pady=3)
        self.create_button(export_box, "Export Excel", self.export_visible_excel).pack(fill=tk.X, pady=3)

        # ---------------------------
        # RIGHT SIDE — STUDENT TABLE + MARKS
        # ---------------------------

        right_frame = tk.Frame(master, bg=BG)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        table_frame = tk.Frame(right_frame, bg=BG)
        table_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("db_id", "student_id", "name", "class", "total", "percentage", "gpa", "grade", "attendance")

        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True)

        for col in cols:
            self.tree.heading(col, text=col.capitalize())
            self.tree.column(col, width=120)

        self.tree.bind("<<TreeviewSelect>>", lambda e: self.on_student_select())

        # MARKS FRAME
        marks_frame = tk.LabelFrame(right_frame, text="Marks Entry",
                                    bg=FRAME_BG, fg=TEXT_COLOR)
        marks_frame.pack(fill=tk.X, pady=8)

        self.lbl_selected_student = tk.Label(marks_frame, text="No student selected",
                                             bg=FRAME_BG, fg=TEXT_COLOR,
                                             font=("Segoe UI", 10, "bold"))
        self.lbl_selected_student.pack(anchor="w")

        self.marks_entries_frame = tk.Frame(marks_frame, bg=FRAME_BG)
        self.marks_entries_frame.pack(fill=tk.X, pady=5)

        self.create_button(marks_frame, "Save Marks", self.save_marks_for_selected).pack(side=tk.LEFT, padx=5, pady=5)
        self.create_button(marks_frame, "View Detailed Marks", self.view_detailed_marks).pack(side=tk.LEFT, padx=5, pady=5)

        # Init
        self.refresh_subjects()
        self.refresh_student_list()
        self.selected_student_db_id = None

    # ---------------------------
    # DARK THEME BUTTON CREATOR
    # ---------------------------

    def create_button(self, parent, text, cmd, w=12):
        btn = tk.Button(parent, text=text, width=w, command=cmd,
                         bg=BTN_BG, fg=TEXT_COLOR, activebackground=BTN_HOVER,
                         activeforeground=TEXT_COLOR, relief="flat", bd=2)
        return btn

    # -----------------------------------------------------------
    # SUBJECT EVENTS
    # -----------------------------------------------------------

    def on_subject_select(self, event):
        sel = self.subj_listbox.curselection()
        if not sel:
            return
        index = sel[0]
        subj_id, name = self.subjects[index]
        self.subject_name_var.set(name)

    # -----------------------------------------------------------
    # SUBJECT CRUD
    # -----------------------------------------------------------

    def add_subject(self):
        name = self.subject_name_var.get().strip()
        if not name:
            messagebox.showerror("Error", "Enter a subject name.")
            return

        add_subject_to_db(name)
        self.subject_name_var.set("")
        self.refresh_subjects()
        self.refresh_student_list()
        messagebox.showinfo("Added", f"Subject '{name}' added successfully.")

    def edit_subject(self):
        sel = self.subj_listbox.curselection()
        if not sel:
            messagebox.showerror("Error", "Select a subject to edit.")
            return

        index = sel[0]
        subj_id, old_name = self.subjects[index]

        new_name = simpledialog.askstring("Edit Subject",
                                          f"Rename '{old_name}' to:",
                                          initialvalue=old_name)
        if not new_name:
            return

        conn = get_conn()
        with closing(conn):
            c = conn.cursor()
            try:
                c.execute("UPDATE subjects SET name=? WHERE id=?", (new_name, subj_id))
                conn.commit()
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", "Subject name already exists.")
                return

        self.refresh_subjects()
        self.refresh_student_list()
        messagebox.showinfo("Updated", f"Subject renamed to '{new_name}'.")

    def delete_subject(self):
        sel = self.subj_listbox.curselection()
        if not sel:
            messagebox.showerror("Error", "Select a subject to delete.")
            return

        index = sel[0]
        subj_id, name = self.subjects[index]

        if not messagebox.askyesno("Delete?", f"Delete subject '{name}'?\nThis removes all related marks."):
            return

        conn = get_conn()
        with closing(conn):
            c = conn.cursor()
            c.execute("DELETE FROM subjects WHERE id=?", (subj_id,))
            c.execute("DELETE FROM marks WHERE subject_db_id=?", (subj_id,))
            conn.commit()

        self.refresh_subjects()
        self.refresh_student_list()
        messagebox.showinfo("Deleted", f"Subject '{name}' removed.")

    # -----------------------------------------------------------
    # REFRESH SUBJECT LIST + MARKS ENTRY UI
    # -----------------------------------------------------------

    def refresh_subjects(self):
        self.subj_listbox.delete(0, tk.END)
        self.subjects = list_subjects_from_db()

        for sid, name in self.subjects:
            self.subj_listbox.insert(tk.END, name)

        for widget in self.marks_entries_frame.winfo_children():
            widget.destroy()

        tk.Label(self.marks_entries_frame, text="Subject",
                 bg=FRAME_BG, fg=TEXT_COLOR).grid(row=0, column=0, sticky="w")
        tk.Label(self.marks_entries_frame, text="Marks",
                 bg=FRAME_BG, fg=TEXT_COLOR).grid(row=0, column=1, sticky="w")

        self.marks_entry_vars = {}

        for i, (sid, name) in enumerate(self.subjects, start=1):
            tk.Label(self.marks_entries_frame, text=name,
                     bg=FRAME_BG, fg=TEXT_COLOR).grid(row=i, column=0, sticky="w")

            var = tk.StringVar()
            e = tk.Entry(self.marks_entries_frame, textvariable=var,
                         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground="white", width=10)
            e.grid(row=i, column=1)
            self.marks_entry_vars[sid] = var

    # -----------------------------------------------------------
    # REFRESH STUDENT LIST
    # -----------------------------------------------------------

    def refresh_student_list(self):
        for r in self.tree.get_children():
            self.tree.delete(r)

        query = self.search_var.get().strip()

        if query:
            students = find_students_by_name(query)
        else:
            students = list_students_from_db()

        for db_id, sid, name, klass in students:
            marks = get_marks_for_student_from_db(db_id)
            total, max_total, percentage, gpa, grade = calculate_total_percentage_gpa_grade(marks)
            attendance = get_attendance_percent_from_db(db_id)

            fg = self.filter_grade_var.get()
            if fg != "All" and grade != fg:
                continue

            self.tree.insert("", tk.END,
                             values=(db_id, sid, name, klass, total, percentage, gpa, grade, attendance))

    # -----------------------------------------------------------
    # STUDENT SELECT EVENT
    # -----------------------------------------------------------

    def on_student_select(self):
        sel = self.tree.selection()
        if not sel:
            return

        item = self.tree.item(sel[0])
        db_id, sid, name, klass = item["values"][:4]

        self.selected_student_db_id = db_id
        self.lbl_selected_student.config(text=f"{name} ({sid})")

        self.sid_var.set(sid)
        self.name_var.set(name)
        self.class_var.set(klass)

        marks = get_marks_for_student_from_db(db_id)
        marks_map = {sname: val for sname, val in marks}

        for subj_id, subj_name in self.subjects:
            if subj_name in marks_map:
                self.marks_entry_vars[subj_id].set(str(marks_map[subj_name]))
            else:
                self.marks_entry_vars[subj_id].set("")

    # -----------------------------------------------------------
    # STUDENT CRUD
    # -----------------------------------------------------------

    def add_student(self):
        sid = self.sid_var.get().strip()
        name = self.name_var.get().strip()
        klass = self.class_var.get().strip()

        if not name:
            messagebox.showerror("Error", "Name is required.")
            return

        if not sid:
            sid = f"S{int(datetime.datetime.now().timestamp())}"

        try:
            add_student_to_db(sid, name, klass)
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", "Student ID already exists.")
            return

        messagebox.showinfo("Added", f"Student '{name}' added.")
        self.refresh_student_list()
        self.clear_form()

    def update_selected_student(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showerror("Error", "Select a student.")
            return

        db_id = self.tree.item(sel[0])["values"][0]

        sid = self.sid_var.get().strip()
        name = self.name_var.get().strip()
        klass = self.class_var.get().strip()

        if not name:
            messagebox.showerror("Error", "Name cannot be empty.")
            return

        try:
            update_student_in_db(db_id, sid, name, klass)
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", "Student ID already exists.")
            return

        messagebox.showinfo("Updated", "Student updated.")
        self.refresh_student_list()
        self.clear_form()

    def delete_selected_student(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showerror("Error", "Select a student.")
            return

        db_id, sid, name = self.tree.item(sel[0])["values"][:3]

        if not messagebox.askyesno("Delete?", f"Delete '{name}'?"):
            return

        delete_student_from_db(db_id)
        messagebox.showinfo("Deleted", f"Student '{name}' removed.")
        self.refresh_student_list()
        self.clear_form()

    # -----------------------------------------------------------
    # SAVE MARKS
    # -----------------------------------------------------------

    def save_marks_for_selected(self):
        if not self.selected_student_db_id:
            messagebox.showerror("Error", "Select a student first.")
            return

        for subj_id, subj_name in self.subjects:
            raw = self.marks_entry_vars[subj_id].get().strip()

            if raw == "":
                continue

            try:
                val = float(raw)
            except:
                messagebox.showerror("Error", f"Invalid marks for {subj_name}.")
                return

            if not 0 <= val <= 100:
                messagebox.showerror("Error", f"Marks for {subj_name} must be 0-100.")
                return

            set_mark_in_db(self.selected_student_db_id, subj_id, val)

        messagebox.showinfo("Saved", "Marks updated.")
        self.refresh_student_list()

    # -----------------------------------------------------------
    # VIEW DETAILED MARKS
    # -----------------------------------------------------------

    def view_detailed_marks(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showerror("Error", "Select a student.")
            return

        db_id, sid, name = self.tree.item(sel[0])["values"][:3]

        marks = get_marks_for_student_from_db(db_id)
        total, max_total, percentage, gpa, grade = calculate_total_percentage_gpa_grade(marks)
        attendance = get_attendance_percent_from_db(db_id)

        msg = f"{name} ({sid})\n\n"
        msg += "Marks:\n"
        for sname, val in marks:
            msg += f"  {sname}: {val}\n"

        msg += f"\nTotal: {total}/{max_total}\n"
        msg += f"Percentage: {percentage}%\n"
        msg += f"GPA: {gpa}\n"
        msg += f"Grade: {grade}\n"
        msg += f"Attendance: {attendance}%"

        messagebox.showinfo("Detailed Marks", msg)

    # -----------------------------------------------------------
    # ATTENDANCE
    # -----------------------------------------------------------

    def mark_present_today(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showerror("Error", "Select a student.")
            return

        db_id, sid, name = self.tree.item(sel[0])["values"][:3]

        today = datetime.date.today().isoformat()
        add_attendance_to_db(db_id, today, True)

        messagebox.showinfo("Marked", f"{name} marked present ({today}).")
        self.refresh_student_list()

    def show_attendance_percent(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showerror("Error", "Select a student.")
            return

        db_id = self.tree.item(sel[0])["values"][0]

        pct = get_attendance_percent_from_db(db_id)
        messagebox.showinfo("Attendance %", f"Attendance: {pct}%")

    # -----------------------------------------------------------
    # EXPORT
    # -----------------------------------------------------------

    def get_visible_rows_for_export(self):
        headers = ["DB_ID", "Student ID", "Name", "Class",
                   "Total", "Percentage", "GPA", "Grade", "Attendance%"]

        rows = [self.tree.item(i)["values"] for i in self.tree.get_children()]

        return rows, headers

    def export_visible_csv(self):
        rows, headers = self.get_visible_rows_for_export()

        if not rows:
            messagebox.showerror("Error", "No data to export.")
            return

        fp = filedialog.asksaveasfilename(defaultextension=".csv",
                                          filetypes=[("CSV files", "*.csv")])
        if not fp:
            return

        with open(fp, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(rows)

        messagebox.showinfo("Exported", f"CSV saved to:\n{fp}")

    def export_visible_excel(self):
        if not PANDAS_AVAILABLE:
            messagebox.showerror("Missing", "Install pandas + openpyxl first.")
            return

        rows, headers = self.get_visible_rows_for_export()
        if not rows:
            messagebox.showerror("Error", "No data to export.")
            return

        fp = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                          filetypes=[("Excel files", "*.xlsx")])
        if not fp:
            return

        df = pd.DataFrame(rows, columns=headers)
        df.to_excel(fp, index=False)

        messagebox.showinfo("Exported", f"Excel saved to:\n{fp}")

    # -----------------------------------------------------------
    # CLEAR FORM & EXIT
    # -----------------------------------------------------------

    def clear_form(self):
        self.sid_var.set("")
        self.name_var.set("")
        self.class_var.set("")
        self.lbl_selected_student.config(text="No student selected")
        self.selected_student_db_id = None

        for v in getattr(self, "marks_entry_vars", {}).values():
            v.set("")

    def clear_search(self):
        self.search_var.set("")
        self.filter_grade_var.set("All")
        self.refresh_student_list()

    def on_close(self):
        if messagebox.askyesno("Exit", "Exit the application?"):
            self.master.destroy()


# -----------------------------------------------------------
# MAIN
# -----------------------------------------------------------

def main():
    init_db()
    root = tk.Tk()
    app = StudentGradingApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
