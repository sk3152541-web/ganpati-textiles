import os
import json
import openpyxl
import time
import datetime
import webbrowser
import pyautogui
import pyperclip
import subprocess
import re
import copy
import calendar
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image as PILImage
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

DATA_FILE = "app_data.json"
PARTIES_FILE = "parties_master.json"
ORDER_FOLDER = r"D:\ORDER PIC 2021-22"

class AutocompleteCombobox(ttk.Combobox):
    def set_completion_list(self, completion_list):
        self._completion_list = sorted(list(set(completion_list)), key=str.lower)
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self['values'] = self._completion_list
        self.bind('<KeyRelease>', self.handle_keyrelease)

    def autocomplete(self, delta=0):
        if not self._hits:
            return
        self.position = (self.position + delta) % len(self._hits)
        self.delete(0, tk.END)
        self.insert(0, self._hits[self.position])
        self.select_range(self.position, tk.END)

    def handle_keyrelease(self, event):
        if event.keysym in ('BackSpace', 'Left', 'Right', 'Up', 'Down', 'Return', 'Tab', 'Shift_L', 'Shift_R', 'Control_L', 'Control_R'):
            return
        typed = self.get()
        if typed == '':
            self['values'] = self._completion_list
        else:
            filtered = [item for item in self._completion_list if typed.lower() in item.lower()]
            self['values'] = filtered

class CalendarPopup:
    def __init__(self, parent, entry_widget):
        self.top = tk.Toplevel(parent)
        self.top.title("Select Date")
        self.top.geometry("280x260")
        self.top.config(bg="#ffffff")
        self.top.resizable(False, False)
        self.top.grab_set()

        self.entry_widget = entry_widget
        
        try:
            current_val = entry_widget.get().strip()
            parts = current_val.split('-')
            if len(parts) == 3:
                self.year = int(parts[0])
                self.month = int(parts[1])
            else:
                now = datetime.datetime.now()
                self.year = now.year
                self.month = now.month
        except:
            now = datetime.datetime.now()
            self.year = now.year
            self.month = now.month

        self.create_widgets()

    def create_widgets(self):
        header_frame = tk.Frame(self.top, bg="#007ACC", pady=8)
        header_frame.pack(fill=tk.X)

        tk.Button(header_frame, text="❮", command=self.prev_month, bg="#007ACC", fg="white", bd=0, font=("Arial", 11, "bold"), activebackground="#005999", activeforeground="white", cursor="hand2").pack(side=tk.LEFT, padx=12)
        
        self.month_year_lbl = tk.Label(header_frame, text="", bg="#007ACC", fg="white", font=("Arial", 11, "bold"))
        self.month_year_lbl.pack(side=tk.LEFT, expand=True)

        tk.Button(header_frame, text="❯", command=self.next_month, bg="#007ACC", fg="white", bd=0, font=("Arial", 11, "bold"), activebackground="#005999", activeforeground="white", cursor="hand2").pack(side=tk.RIGHT, padx=12)

        days_frame = tk.Frame(self.top, bg="#f0f4f8", pady=4)
        days_frame.pack(fill=tk.X, padx=10, pady=5)

        days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        for i, d in enumerate(days):
            tk.Label(days_frame, text=d, bg="#f0f4f8", fg="#1f4e78", font=("Arial", 8, "bold"), width=4).grid(row=0, column=i, padx=1)

        self.cal_frame = tk.Frame(self.top, bg="#ffffff")
        self.cal_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=2)

        self.build_calendar()

    def build_calendar(self):
        for widget in self.cal_frame.winfo_children():
            widget.destroy()

        self.month_year_lbl.config(text=f"{calendar.month_name[self.month]} {self.year}")

        today = datetime.date.today()
        cal_matrix = calendar.monthcalendar(self.year, self.month)
        
        for r_idx, week in enumerate(cal_matrix):
            for c_idx, day in enumerate(week):
                if day == 0:
                    tk.Label(self.cal_frame, text="", bg="#ffffff", width=4, height=1).grid(row=r_idx, column=c_idx, padx=1, pady=1)
                else:
                    is_today = (self.year == today.year and self.month == today.month and day == today.day)
                    
                    btn_bg = "#e6f2ff" if is_today else "#f8f9fa"
                    btn_fg = "#007ACC" if is_today else "#333333"
                    font_weight = "bold" if is_today else "normal"

                    btn = tk.Button(self.cal_frame, text=str(day), width=3, height=1, bg=btn_bg, fg=btn_fg,
                                    font=("Arial", 9, font_weight), relief=tk.FLAT,
                                    activebackground="#007ACC", activeforeground="white",
                                    cursor="hand2",
                                    command=lambda d=day: self.select_date(d))
                    btn.grid(row=r_idx, column=c_idx, padx=2, pady=2)

    def prev_month(self):
        self.month -= 1
        if self.month < 1:
            self.month = 12
            self.year -= 1
        self.build_calendar()

    def next_month(self):
        self.month += 1
        if self.month > 12:
            self.month = 1
            self.year += 1
        self.build_calendar()

    def select_date(self, day):
        selected_date = f"{self.year:04d}-{self.month:02d}-{day:02d}"
        self.entry_widget.delete(0, tk.END)
        self.entry_widget.insert(0, selected_date)
        self.top.destroy()

class LoginWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Ganpati Textiles - Login")
        self.root.geometry("380x240")
        self.root.config(bg="#1e1e1e")
        self.root.resizable(False, False)

        self.root.eval('tk::PlaceWindow . center')

        main_frame = tk.Frame(root, bg="#2d2d2d", bd=2, relief=tk.GROOVE)
        main_frame.place(x=10, y=10, width=360, height=220)

        tk.Label(main_frame, text="Ganpati Textiles - Login", font=("Arial", 11, "bold"), bg="#2d2d2d", fg="#4da6ff").pack(pady=12)

        form_frame = tk.Frame(main_frame, bg="#2d2d2d")
        form_frame.pack(padx=20, pady=5, fill=tk.X)

        tk.Label(form_frame, text="USER NAME:", font=("Arial", 9, "bold"), bg="#2d2d2d", fg="#ffffff").grid(row=0, column=0, sticky="w", pady=5)
        
        self.user_var = tk.StringVar()
        self.user_entry = ttk.Combobox(form_frame, textvariable=self.user_var, font=("Arial", 10), width=20, state="normal")
        self.user_entry['values'] = ("admin", "ganpati", "manager")
        self.user_entry.set("") 
        self.user_entry.grid(row=0, column=1, padx=10, pady=5)
        self.user_entry.focus_set()
        self.user_entry.bind("<Return>", lambda event: self.pass_entry.focus_set())

        tk.Label(form_frame, text="PASSWORD:", font=("Arial", 9, "bold"), bg="#2d2d2d", fg="#ffffff").grid(row=1, column=0, sticky="w", pady=5)
        self.pass_entry = tk.Entry(form_frame, show="*", font=("Arial", 10), width=22)
        self.pass_entry.grid(row=1, column=1, padx=10, pady=5)
        self.pass_entry.bind("<Return>", lambda event: self.verify_login())

        btn_frame = tk.Frame(main_frame, bg="#2d2d2d")
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="Login", command=self.verify_login, bg="#1f4e78", fg="white", font=("Arial", 9, "bold"), width=10, relief=tk.RAISED).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Exit", command=root.destroy, bg="#a60000", fg="white", font=("Arial", 9, "bold"), width=10, relief=tk.RAISED).pack(side=tk.LEFT, padx=10)

    def verify_login(self):
        correct_user = "admin"
        correct_pass = "admin123"

        entered_user = self.user_var.get().strip()
        entered_pass = self.pass_entry.get().strip()

        if entered_user == correct_user and entered_pass == correct_pass:
            self.root.destroy()
            main_root = tk.Tk()
            app = GanpatiSyncDashboard(main_root)
            main_root.mainloop()
        else:
            messagebox.showerror("Access Denied", "Galat Username ya Password! Kripya dobara koshish karein.")
            self.pass_entry.delete(0, tk.END)
            self.user_entry.focus_set()

class GanpatiSyncDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("Ganpati Textiles - Independent App Dashboard")
        self.root.state('zoomed')
        self.root.config(bg="#f4f4f4")

        self.undo_stack = []
        self.redo_stack = []
        self.is_call_to_party_filter_active = False

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", font=("Arial", 9, "bold"), background="#1f4e78", foreground="white", relief="flat")
        style.configure("Treeview", font=("Arial", 9), rowheight=25, fieldbackground="white")
        style.map("Treeview", background=[('selected', '#007ACC')], foreground=[('selected', 'white')])

        title_lbl = tk.Label(root, text="Ganpati Textiles - Independent Dashboard with Multi-Status Filtering, Sorting & Live Remark Preview", font=("Arial", 13, "bold"), bg="#1f4e78", fg="white", pady=6)
        title_lbl.pack(fill=tk.X)

        # Main Toolbar Frame
        btn_frame = tk.Frame(root, bg="#f4f4f4", pady=4)
        btn_frame.pack(fill=tk.X, padx=10)

        # Row 1: Action Buttons & WhatsApp Button
        row1 = tk.Frame(btn_frame, bg="#f4f4f4")
        row1.pack(fill=tk.X, pady=2)

        tk.Button(row1, text="📂 Reload", command=self.load_data, bg="#007ACC", fg="white", font=("Arial", 8, "bold"), padx=3, pady=2).pack(side=tk.LEFT, padx=1)
        tk.Button(row1, text="📊 Export", command=self.export_to_excel, bg="#00897B", fg="white", font=("Arial", 8, "bold"), padx=3, pady=2).pack(side=tk.LEFT, padx=1)
        tk.Button(row1, text="📥 Import", command=self.import_from_excel, bg="#673AB7", fg="white", font=("Arial", 8, "bold"), padx=3, pady=2).pack(side=tk.LEFT, padx=1)
        tk.Button(row1, text="🔍 Call Filter", command=self.toggle_call_to_party_filter, bg="#00695C", fg="white", font=("Arial", 8, "bold"), padx=3, pady=2).pack(side=tk.LEFT, padx=1)
        tk.Button(row1, text="👁 Hide", command=self.hide_selected_orders, bg="#455A64", fg="white", font=("Arial", 8, "bold"), padx=3, pady=2).pack(side=tk.LEFT, padx=1)
        tk.Button(row1, text="👁️‍🗨️ Unhide", command=self.unhide_all_orders, bg="#37474F", fg="white", font=("Arial", 8, "bold"), padx=3, pady=2).pack(side=tk.LEFT, padx=1)
        tk.Button(row1, text="✅ Check", command=self.check_all, bg="#4CAF50", fg="white", font=("Arial", 8, "bold"), padx=3, pady=2).pack(side=tk.LEFT, padx=1)
        tk.Button(row1, text="❌ Uncheck", command=self.uncheck_all, bg="#f44336", fg="white", font=("Arial", 8, "bold"), padx=3, pady=2).pack(side=tk.LEFT, padx=1)
        tk.Button(row1, text="➕ Add", command=self.open_add_window, bg="#FF9800", fg="white", font=("Arial", 8, "bold"), padx=3, pady=2).pack(side=tk.LEFT, padx=1)
        tk.Button(row1, text="✏️ Modify", command=self.modify_selected_order, bg="#0288D1", fg="white", font=("Arial", 8, "bold"), padx=3, pady=2).pack(side=tk.LEFT, padx=1)
        tk.Button(row1, text="💬 Reply", command=self.edit_remark_popup, bg="#607D8B", fg="white", font=("Arial", 8, "bold"), padx=3, pady=2).pack(side=tk.LEFT, padx=1)
        tk.Button(row1, text="🗑 Delete", command=self.delete_selected_orders, bg="#D32F2F", fg="white", font=("Arial", 8, "bold"), padx=3, pady=2).pack(side=tk.LEFT, padx=1)
        tk.Button(row1, text="↩ Undo", command=self.undo_action, bg="#78909C", fg="white", font=("Arial", 8, "bold"), padx=3, pady=2).pack(side=tk.LEFT, padx=1)
        tk.Button(row1, text="↪ Redo", command=self.redo_action, bg="#78909C", fg="white", font=("Arial", 8, "bold"), padx=3, pady=2).pack(side=tk.LEFT, padx=1)
        tk.Button(row1, text="🚀 WhatsApp", command=self.send_whatsapp_messages, bg="#2E7D32", fg="white", font=("Arial", 8, "bold"), padx=4, pady=2).pack(side=tk.LEFT, padx=6)

        # Row 2: Status Filters & Search Box
        row2 = tk.Frame(btn_frame, bg="#f4f4f4")
        row2.pack(fill=tk.X, pady=2)

        status_frame = tk.Frame(row2, bg="#e0e0e0", bd=1, relief=tk.GROOVE, padx=1, pady=1)
        status_frame.pack(side=tk.LEFT, padx=1)
        
        tk.Label(status_frame, text="Filter:", bg="#e0e0e0", font=("Arial", 8, "bold")).pack(side=tk.LEFT, padx=1)

        self.status_vars = {
            "All": tk.BooleanVar(value=True),
            "Pending": tk.BooleanVar(value=False),
            "Expiry Soon": tk.BooleanVar(value=False),
            "Expired": tk.BooleanVar(value=False)
        }

        def on_status_toggle(changed_key):
            self.is_call_to_party_filter_active = False
            if changed_key == "All":
                if self.status_vars["All"].get():
                    for k in ["Pending", "Expiry Soon", "Expired"]:
                        self.status_vars[k].set(False)
            else:
                if self.status_vars[changed_key].get():
                    self.status_vars["All"].set(False)
                else:
                    if not any(self.status_vars[k].get() for k in ["Pending", "Expiry Soon", "Expired"]):
                        self.status_vars["All"].set(True)
            self.apply_filter()

        for st_name in ["All", "Pending", "Expiry Soon", "Expired"]:
            cb = tk.Checkbutton(status_frame, text=st_name, variable=self.status_vars[st_name],
                                command=lambda k=st_name: on_status_toggle(k),
                                bg="#e0e0e0", font=("Arial", 8, "bold"), selectcolor="#ffffff")
            cb.pack(side=tk.LEFT, padx=1)

        search_frame = tk.Frame(row2, bg="#f4f4f4")
        search_frame.pack(side=tk.LEFT, padx=10)
        tk.Label(search_frame, text="🔍 Search:", bg="#f4f4f4", font=("Arial", 8, "bold")).pack(side=tk.LEFT, padx=1)
        
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.apply_filter())
        tk.Entry(search_frame, textvariable=self.search_var, width=15, font=("Arial", 8)).pack(side=tk.LEFT, padx=1)

        table_container = tk.Frame(root, bg="white")
        table_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        self.columns_def = ("select", "order_no", "date", "party", "design", "shirt", "dly_date", "phone", "status", "remarks")
        
        self.tree = ttk.Treeview(table_container, columns=self.columns_def, show="headings", selectmode="extended")
        
        headers_text = [
            ("select", "SELECT ▾", 45, "center"),
            ("order_no", "ORDER NO ▾", 75, "center"),
            ("date", "DATE ▾", 85, "center"),
            ("party", "PARTY NAME ▾", 160, "w"),
            ("design", "DESIGN NO ▾", 85, "center"),
            ("shirt", "SHIRT COLOUR ▾", 90, "center"),
            ("dly_date", "DLY. DATE ▾", 85, "center"),
            ("phone", "PHONE ▾", 85, "center"),
            ("status", "STATUS ▾", 85, "center"),
            ("remarks", "PARTY REPLY / REMARKS ▾", 320, "w")
        ]

        for col_id, col_name, width, anchor in headers_text:
            self.tree.heading(col_id, text=col_name, anchor=anchor, command=lambda c=col_id: self.show_sort_menu(c))
            self.tree.column(col_id, width=width, anchor=anchor, stretch=True)

        scrollbar = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        preview_frame = tk.Frame(root, bg="#e8f4f8", bd=1, relief=tk.SOLID)
        preview_frame.pack(fill=tk.X, padx=10, pady=4)
        
        tk.Label(preview_frame, text="💬 Live Remark Preview (Click any row to read full reply instantly):", font=("Arial", 8, "bold"), bg="#e8f4f8", fg="#1f4e78").pack(anchor="w", padx=8, pady=2)
        
        self.preview_lbl = tk.Label(preview_frame, text="[Select any row above to view full remarks here...]", font=("Arial", 9, "bold"), bg="#e8f4f8", fg="#b26b00", anchor="w", justify=tk.LEFT, wraplength=1300)
        self.preview_lbl.pack(fill=tk.X, padx=8, pady=3)

        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)

        self.all_data = []
        self.load_data()

    def get_master_parties(self):
        parties = set()
        if os.path.exists(PARTIES_FILE):
            try:
                with open(PARTIES_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        parties.update(data)
            except:
                pass
        
        for r in self.all_data:
            p = str(r.get("party", "")).strip()
            if p:
                parties.add(p)
                
        return sorted(list(parties), key=str.lower)

    def save_party_to_master(self, party_name):
        party_name = str(party_name).strip()
        if not party_name:
            return
        parties = set()
        if os.path.exists(PARTIES_FILE):
            try:
                with open(PARTIES_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        parties.update(data)
            except:
                pass
        parties.add(party_name)
        try:
            with open(PARTIES_FILE, "w", encoding="utf-8") as f:
                json.dump(sorted(list(parties), key=str.lower), f, ensure_ascii=False, indent=4)
        except Exception:
            pass

    def save_state_to_undo(self):
        self.undo_stack.append(copy.deepcopy(self.all_data))
        self.redo_stack.clear()

    def undo_action(self):
        if not self.undo_stack:
            messagebox.showinfo("Undo", "Undo karne ke liye kuch nahi bacha hai!")
            return
        self.redo_stack.append(copy.deepcopy(self.all_data))
        self.all_data = self.undo_stack.pop()
        self.save_to_local_storage()
        self.apply_filter()

    def redo_action(self):
        if not self.redo_stack:
            messagebox.showinfo("Redo", "Redo karne ke liye kuch nahi bacha hai!")
            return
        self.undo_stack.append(copy.deepcopy(self.all_data))
        self.all_data = self.redo_stack.pop()
        self.save_to_local_storage()
        self.apply_filter()

    def save_to_local_storage(self):
        try:
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(self.all_data, f, ensure_ascii=False, indent=4)
            for r in self.all_data:
                self.save_party_to_master(r.get("party", ""))
        except Exception as e:
            messagebox.showerror("Error", f"Storage save error: {e}")

    def load_data(self):
        if not os.path.exists(DATA_FILE):
            self.all_data = []
            return

        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                self.all_data = json.load(f)
            
            for row in self.all_data:
                if "hidden" not in row:
                    row["hidden"] = False
                party_name = row.get("party", "")
                if party_name:
                    self.save_party_to_master(party_name)

                dly_date_str = row.get("dly_date", "")
                status = "Pending"
                if dly_date_str:
                    try:
                        d_date = datetime.datetime.strptime(dly_date_str, "%Y-%m-%d").date()
                        diff = (d_date - datetime.date.today()).days
                        if diff < 0: status = "Expired"
                        elif diff <= 7: status = "Expiry Soon"
                    except:
                        pass
                row["status"] = status

            self.apply_filter()
        except Exception as e:
            messagebox.showerror("Error", f"Data load error: {e}")

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Export Data to Excel")
        if file_path:
            try:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Orders Export"
                headers = ["ORDER NO", "DATE", "PARTY NAME", "DESIGN NO", "SHIRT COLOUR", "DLY. DATE", "PHONE", "STATUS", "REMARKS"]
                ws.append(headers)
                for row in self.all_data:
                    if not row.get("hidden", False):
                        ws.append([
                            row["order_no"], row["date"], row["party"],
                            row["design"], row["shirt"], row["dly_date"],
                            row["phone"], row["status"], row["remarks"]
                        ])
                wb.save(file_path)
                messagebox.showinfo("Success", f"Data successfully exported to:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {e}")

    def import_from_excel(self):
        file_path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Import Data from Excel")
        if file_path:
            try:
                wb = openpyxl.load_workbook(file_path, data_only=True)
                ws = wb.active
                new_imported_data = []

                for row in range(2, ws.max_row + 1):
                    order_no = ws.cell(row=row, column=1).value
                    if order_no:
                        date = ws.cell(row=row, column=2).value
                        party = ws.cell(row=row, column=3).value
                        design = ws.cell(row=row, column=4).value
                        shirt = ws.cell(row=row, column=5).value
                        dly_date = ws.cell(row=row, column=6).value
                        phone = ws.cell(row=row, column=7).value
                        status = ws.cell(row=row, column=8).value
                        remarks = ws.cell(row=row, column=9).value

                        date_str = str(date).split()[0] if date else ""
                        dly_date_str = str(dly_date).split()[0] if dly_date else ""

                        party_str = str(party or "").strip()
                        if party_str:
                            self.save_party_to_master(party_str)

                        calc_status = "Pending"
                        if dly_date_str:
                            try:
                                d_date = datetime.datetime.strptime(dly_date_str, "%Y-%m-%d").date()
                                diff = (d_date - datetime.date.today()).days
                                if diff < 0: calc_status = "Expired"
                                elif diff <= 7: calc_status = "Expiry Soon"
                            except:
                                pass

                        new_imported_data.append({
                            "select": False,
                            "hidden": False,
                            "order_no": str(order_no or "").strip(),
                            "date": date_str,
                            "party": party_str,
                            "design": str(design or ""),
                            "shirt": str(shirt or ""),
                            "dly_date": dly_date_str,
                            "phone": str(phone or "").strip(),
                            "status": calc_status,
                            "remarks": str(remarks or "").strip()
                        })

                if new_imported_data:
                    self.save_state_to_undo()
                    self.all_data = new_imported_data
                    self.save_to_local_storage()
                    self.apply_filter()
                    messagebox.showinfo("Success", "Excel file ka data successfully import aur save ho gaya!")
            except Exception as e:
                messagebox.showerror("Error", f"Import failed: {e}")

    def toggle_call_to_party_filter(self):
        self.is_call_to_party_filter_active = not self.is_call_to_party_filter_active
        if self.is_call_to_party_filter_active:
            messagebox.showinfo("Filter Active", "Call To Party Filter on ho gaya hai. Dobara click karne par normal view aa jayega.")
        else:
            messagebox.showinfo("Filter Off", "Call To Party Filter off kar diya gaya hai.")
        self.apply_filter()

    def hide_selected_orders(self):
        selected = self.get_selected_row_objects()
        if not selected:
            messagebox.showwarning("Warning", "Kripya hide karne ke liye kam se kam ek order select (☑) karein!")
            return
        
        self.save_state_to_undo()
        for r in selected:
            r["hidden"] = True
            r["select"] = False
        self.save_to_local_storage()
        self.apply_filter()
        messagebox.showinfo("Hidden", f"Selected {len(selected)} orders hide kar diye gaye hain!")

    def unhide_all_orders(self):
        hidden_count = sum(1 for r in self.all_data if r.get("hidden", False))
        if hidden_count == 0:
            messagebox.showinfo("Unhide", "Koi bhi order hidden nahi hai!")
            return
        
        self.save_state_to_undo()
        for r in self.all_data:
            r["hidden"] = False
        self.save_to_local_storage()
        self.apply_filter()
        messagebox.showinfo("Unhide", f"Sabhi {hidden_count} hidden orders wapas unhide kar diye gaye hain!")

    def show_sort_menu(self, col):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Sort A to Z", command=lambda: self.sort_column(col, False))
        menu.add_command(label="Sort Z to A", command=lambda: self.sort_column(col, True))
        
        try:
            x = self.root.winfo_pointerx()
            y = self.root.winfo_pointery()
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def sort_column(self, col, reverse):
        key_mapping = {
            "select": "select",
            "order_no": "order_no",
            "date": "date",
            "party": "party",
            "design": "design",
            "shirt": "shirt",
            "dly_date": "dly_date",
            "phone": "phone",
            "status": "status",
            "remarks": "remarks"
        }
        field = key_mapping.get(col)
        if field:
            try:
                self.all_data.sort(key=lambda x: str(x.get(field, "")).lower(), reverse=reverse)
            except Exception:
                pass
            self.apply_filter()

    def apply_filter(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        search_text = self.search_var.get().strip().lower() if hasattr(self, 'search_var') else ""
        
        is_all = self.status_vars["All"].get()
        active_statuses = [st for st in ["Pending", "Expiry Soon", "Expired"] if self.status_vars[st].get()]

        for idx, row in enumerate(self.all_data):
            if row.get("hidden", False):
                continue

            if self.is_call_to_party_filter_active:
                if "CALL TO PARTY" not in str(row["remarks"]).upper():
                    continue
            else:
                if not is_all and active_statuses and row["status"] not in active_statuses:
                    continue

            if search_text:
                match = (
                    search_text in row["order_no"].lower() or
                    search_text in row["party"].lower() or
                    search_text in row["design"].lower() or
                    search_text in row["shirt"].lower() or
                    search_text in row["phone"].lower() or
                    search_text in row["remarks"].lower()
                )
                if not match:
                    continue

            sel_icon = "☑" if row["select"] else "☐"
            values = (
                sel_icon, row["order_no"], row["date"], row["party"],
                row["design"], row["shirt"], row["dly_date"], row["phone"],
                row["status"], row["remarks"]
            )
            self.tree.insert("", tk.END, values=values, tags=(row["status"], str(idx)))

        self.tree.tag_configure("Expired", background="#ffcccc", foreground="#990000")
        self.tree.tag_configure("Expiry Soon", background="#fff2cc", foreground="#b26b00")
        self.tree.tag_configure("Pending", background="#e8f5e9", foreground="#274e13")

    def on_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            item_id = self.tree.identify_row(event.y)
            if column == "#1" and item_id:
                try:
                    tags = self.tree.item(item_id, "tags")
                    if tags and len(tags) > 1:
                        data_idx = int(tags[1])
                        self.all_data[data_idx]["select"] = not self.all_data[data_idx]["select"]
                        
                        sel_icon = "☑" if self.all_data[data_idx]["select"] else "☐"
                        current_values = list(self.tree.item(item_id, "values"))
                        current_values[0] = sel_icon
                        self.tree.item(item_id, values=current_values)
                except Exception:
                    pass

    def on_row_select(self, event):
        selected_items = self.tree.selection()
        if selected_items:
            item_id = selected_items[0]
            values = self.tree.item(item_id, "values")
            if values:
                ord_no = values[1]
                party_name = values[3]
                for r in self.all_data:
                    if r["order_no"] == ord_no and r["party"] == party_name:
                        remark_text = r["remarks"] if r["remarks"] else "[No remarks available]"
                        self.preview_lbl.config(text=f"Order: {ord_no} | Party: {party_name}\nReply: {remark_text}", fg="#000000")
                        break

    def on_tree_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            if column == "#1":
                return
        
        item_id = self.tree.identify_row(event.y)
        if item_id:
            try:
                tags = self.tree.item(item_id, "tags")
                if tags and len(tags) > 1:
                    data_idx = int(tags[1])
                    self.edit_remark_popup_for_row(self.all_data[data_idx])
            except Exception:
                pass

    def check_all(self):
        for item in self.tree.get_children():
            tags = self.tree.item(item, "tags")
            if tags and len(tags) > 1:
                data_idx = int(tags[1])
                self.all_data[data_idx]["select"] = True
        self.apply_filter()

    def uncheck_all(self):
        for item in self.tree.get_children():
            tags = self.tree.item(item, "tags")
            if tags and len(tags) > 1:
                data_idx = int(tags[1])
                self.all_data[data_idx]["select"] = False
        self.apply_filter()

    def get_selected_row_objects(self):
        return [r for r in self.all_data if r["select"] and not r.get("hidden", False)]

    def delete_selected_orders(self):
        selected = self.get_selected_row_objects()
        if not selected:
            messagebox.showwarning("Warning", "Kripya delete karne ke liye kam se kam ek order select (☑) karein!")
            return
        
        if messagebox.askyesno("Confirm Delete", f"Kya aap sach mein selected {len(selected)} orders ko delete karna chahte hain?"):
            self.save_state_to_undo()
            self.all_data = [r for r in self.all_data if not r["select"]]
            self.save_to_local_storage()
            self.apply_filter()
            messagebox.showinfo("Deleted", "Selected orders successfully delete kar diye gaye hain!")

    def modify_selected_order(self):
        selected = self.get_selected_row_objects()
        if not selected:
            messagebox.showwarning("Warning", "Kripya modify karne ke liye kam se kam ek order select (☑) karein!")
            return
        if len(selected) > 1:
            messagebox.showwarning("Warning", "Kripya ek waqt mein sirf ek hi order modify karein!")
            return
        
        target_row = selected[0]
        
        mod_win = tk.Toplevel(self.root)
        mod_win.title(f"Modify Order - {target_row['order_no']}")
        mod_win.geometry("380x560")
        mod_win.config(bg="#f4f4f4")

        tk.Label(mod_win, text="Modify Order Details", font=("Arial", 11, "bold"), bg="#f4f4f4", fg="#1f4e78").pack(pady=10)
        
        fields = [
            ("Order No:", target_row["order_no"]),
            ("Date (YYYY-MM-DD):", target_row["date"]),
            ("Party Name:", target_row["party"]),
            ("Design No:", target_row["design"]),
            ("Shirt Colour:", target_row["shirt"]),
            ("Delivery Date (YYYY-MM-DD):", target_row["dly_date"]),
            ("Phone No:", target_row["phone"]),
            ("Remarks:", target_row["remarks"])
        ]
        
        widgets = []
        for idx, (label_text, default_val) in enumerate(fields):
            tk.Label(mod_win, text=label_text, bg="#f4f4f4", font=("Arial", 9)).pack(anchor="w", padx=20)
            
            if idx == 2:  # Party Name Field as Autocomplete Combobox
                widget = AutocompleteCombobox(mod_win, font=("Arial", 9), width=30)
                widget.set_completion_list(self.get_master_parties())
                widget.pack(padx=20, pady=2)
                widget.set(str(default_val or ""))
                widgets.append(widget)
            elif idx in (1, 5):  # Date fields with Calendar button
                f_date = tk.Frame(mod_win, bg="#f4f4f4")
                f_date.pack(padx=20, pady=2, anchor="w")
                widget = tk.Entry(f_date, width=24, font=("Arial", 9))
                widget.pack(side=tk.LEFT)
                widget.insert(0, str(default_val or ""))
                
                cal_btn = tk.Button(f_date, text="📅", command=lambda w=widget: CalendarPopup(mod_win, w), bg="#007ACC", fg="white", font=("Arial", 8), cursor="hand2")
                cal_btn.pack(side=tk.LEFT, padx=5)
                widgets.append(widget)
            else:
                widget = tk.Entry(mod_win, width=32, font=("Arial", 9))
                widget.pack(padx=20, pady=2)
                widget.insert(0, str(default_val or ""))
                widgets.append(widget)

        widgets[0].focus_set()

        def save_modifications():
            vals = [w.get().strip() for w in widgets]
            if not vals[0] or not vals[2]:
                messagebox.showerror("Error", "Order No aur Party Name zaroori hai!")
                return
            
            self.save_party_to_master(vals[2])
            self.save_state_to_undo()
            target_row["order_no"] = vals[0]
            target_row["date"] = vals[1]
            target_row["party"] = vals[2]
            target_row["design"] = vals[3]
            target_row["shirt"] = vals[4]
            target_row["dly_date"] = vals[5]
            target_row["phone"] = vals[6]
            target_row["remarks"] = vals[7]
            
            dly_date_str = vals[5]
            status = "Pending"
            if dly_date_str:
                try:
                    d_date = datetime.datetime.strptime(dly_date_str, "%Y-%m-%d").date()
                    diff = (d_date - datetime.date.today()).days
                    if diff < 0: status = "Expired"
                    elif diff <= 7: status = "Expiry Soon"
                except:
                    pass
            target_row["status"] = status

            self.save_to_local_storage()
            self.apply_filter()
            mod_win.destroy()
            messagebox.showinfo("Success", "Order successfully modify kar diya gaya hai!")

        def handle_key_navigation(event, current_idx):
            if event.keysym == 'Return':
                if current_idx < len(widgets) - 1:
                    widgets[current_idx + 1].focus_set()
                else:
                    save_modifications()
                return "break"
            elif event.keysym == 'Down':
                if current_idx < len(widgets) - 1:
                    widgets[current_idx + 1].focus_set()
                return "break"
            elif event.keysym == 'Up':
                if current_idx > 0:
                    widgets[current_idx - 1].focus_set()
                return "break"

        for idx, w in enumerate(widgets):
            w.bind("<Key>", lambda event, i=idx: handle_key_navigation(event, i))

        tk.Button(mod_win, text="Update Order", command=save_modifications, bg="#0288D1", fg="white", font=("Arial", 9, "bold"), padx=10, pady=4).pack(pady=15)

    def edit_remark_popup(self):
        selected = self.get_selected_row_objects()
        if not selected:
            messagebox.showwarning("Warning", "Kripya koi order select karein!")
            return
        self.edit_remark_popup_for_row(selected[0])

    def edit_remark_popup_for_row(self, target_row):
        popup = tk.Toplevel(self.root)
        popup.title(f"Full Party Reply - {target_row['party']}")
        popup.geometry("450x300")
        popup.config(bg="#f4f4f4")

        tk.Label(popup, text=f"Order No: {target_row['order_no']} | Party: {target_row['party']}", font=("Arial", 10, "bold"), bg="#f4f4f4", fg="#1f4e78").pack(pady=10)
        tk.Label(popup, text="Party Reply (Aap yahan text read aur edit kar sakte hain):", font=("Arial", 9), bg="#f4f4f4").pack(anchor="w", padx=20)

        txt_box = tk.Text(popup, wrap=tk.WORD, font=("Arial", 10), width=48, height=8)
        txt_box.pack(padx=20, pady=5, fill=tk.BOTH, expand=True)
        txt_box.insert("1.0", target_row["remarks"])

        def save_wrapped_remark():
            new_remark = txt_box.get("1.0", tk.END).strip()
            if new_remark != target_row["remarks"]:
                self.save_state_to_undo()
                target_row["remarks"] = new_remark
                self.save_to_local_storage()
                self.apply_filter()
            popup.destroy()

        tk.Button(popup, text="Save Reply", command=save_wrapped_remark, bg="#2E7D32", fg="white", font=("Arial", 9, "bold"), padx=10, pady=4).pack(pady=10)

    def open_add_window(self):
        add_win = tk.Toplevel(self.root)
        add_win.title("Add New Order")
        add_win.geometry("380x560")
        add_win.config(bg="#f4f4f4")

        tk.Label(add_win, text="Add New Order Entry", font=("Arial", 11, "bold"), bg="#f4f4f4", fg="#1f4e78").pack(pady=10)
        
        fields = [
            ("Order No:", ""),
            ("Date (YYYY-MM-DD):", datetime.date.today().strftime("%Y-%m-%d")),
            ("Party Name:", ""),
            ("Design No:", ""),
            ("Shirt Colour:", ""),
            ("Delivery Date (YYYY-MM-DD):", ""),
            ("Phone No:", ""),
            ("Remarks:", "")
        ]
        
        widgets = []
        for idx, (label_text, default_val) in enumerate(fields):
            tk.Label(add_win, text=label_text, bg="#f4f4f4", font=("Arial", 9)).pack(anchor="w", padx=20)
            
            if idx == 2:  # Party Name Field as Autocomplete Combobox
                widget = AutocompleteCombobox(add_win, font=("Arial", 9), width=30)
                widget.set_completion_list(self.get_master_parties())
                widget.pack(padx=20, pady=2)
                widget.set(str(default_val))
                widgets.append(widget)
            elif idx in (1, 5):  # Date fields with Calendar button
                f_date = tk.Frame(add_win, bg="#f4f4f4")
                f_date.pack(padx=20, pady=2, anchor="w")
                widget = tk.Entry(f_date, width=24, font=("Arial", 9))
                widget.pack(side=tk.LEFT)
                widget.insert(0, str(default_val))
                
                cal_btn = tk.Button(f_date, text="📅", command=lambda w=widget: CalendarPopup(add_win, w), bg="#007ACC", fg="white", font=("Arial", 8), cursor="hand2")
                cal_btn.pack(side=tk.LEFT, padx=5)
                widgets.append(widget)
            else:
                widget = tk.Entry(add_win, width=32, font=("Arial", 9))
                widget.pack(padx=20, pady=2)
                widget.insert(0, str(default_val))
                widgets.append(widget)

        widgets[0].focus_set()

        def save_new():
            vals = [w.get().strip() for w in widgets]
            if not vals[0] or not vals[2]:
                messagebox.showerror("Error", "Order No aur Party Name zaroori hai!")
                return
            
            self.save_party_to_master(vals[2])
            self.save_state_to_undo()
            dly_date_str = vals[5]
            status = "Pending"
            if dly_date_str:
                try:
                    d_date = datetime.datetime.strptime(dly_date_str, "%Y-%m-%d").date()
                    diff = (d_date - datetime.date.today()).days
                    if diff < 0: status = "Expired"
                    elif diff <= 7: status = "Expiry Soon"
                except:
                    pass

            new_row = {
                "select": False,
                "hidden": False,
                "order_no": vals[0],
                "date": vals[1],
                "party": vals[2],
                "design": vals[3],
                "shirt": vals[4],
                "dly_date": vals[5],
                "phone": vals[6],
                "status": status,
                "remarks": vals[7]
            }

            self.all_data.append(new_row)
            self.save_to_local_storage()
            self.apply_filter()
            add_win.destroy()
            messagebox.showinfo("Success", "Naya order successfully add ho gaya hai!")

        def handle_key_navigation(event, current_idx):
            if event.keysym == 'Return':
                if current_idx < len(widgets) - 1:
                    widgets[current_idx + 1].focus_set()
                else:
                    save_new()
                return "break"
            elif event.keysym == 'Down':
                if current_idx < len(widgets) - 1:
                    widgets[current_idx + 1].focus_set()
                return "break"
            elif event.keysym == 'Up':
                if current_idx > 0:
                    widgets[current_idx - 1].focus_set()
                return "break"

        for idx, w in enumerate(widgets):
            w.bind("<Key>", lambda event, i=idx: handle_key_navigation(event, i))

        tk.Button(add_win, text="Save Order", command=save_new, bg="#2E7D32", fg="white", font=("Arial", 9, "bold"), padx=10, pady=4).pack(pady=15)

    def send_whatsapp_messages(self):
        selected_items = self.get_selected_row_objects()
        if not selected_items:
            messagebox.showwarning("Warning", "Kripya kam se kam ek order select (☑) karein!")
            return

        parties_dict = {}
        for item in selected_items:
            p_name = item["party"]
            if p_name not in parties_dict:
                parties_dict[p_name] = {"orders": []}
            parties_dict[p_name]["orders"].append(item)

        success_count = 0
        pyautogui.PAUSE = 1.0
        webbrowser.open("https://web.whatsapp.com/")
        time.sleep(12)

        for party_name, data in parties_dict.items():
            if not os.path.exists(ORDER_FOLDER):
                os.makedirs(ORDER_FOLDER, exist_ok=True)

            safe_party_name = "".join([c if c.isalnum() else "_" for c in party_name])
            pdf_filename = f"List_{safe_party_name}.pdf"
            pdf_path = os.path.join(ORDER_FOLDER, pdf_filename)

            doc = SimpleDocTemplate(pdf_path, pagesize=letter, leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=30)
            elements = []
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, alignment=1, textColor=colors.HexColor('#1f4e78'), spaceAfter=10)
            cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=8, alignment=1)
            header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontSize=9, alignment=1, fontName='Helvetica-Bold', textColor=colors.whitesmoke)

            elements.append(Paragraph("<b>GANPATI TEXTILES</b>", title_style))
            elements.append(Paragraph(f"<b>PENDING LIST FOR: {party_name}</b>", ParagraphStyle('SubTitle', parent=styles['Normal'], alignment=1, fontSize=11, spaceAfter=15)))
            elements.append(Spacer(1, 5))

            table_data = [[
                Paragraph("ORDER NO", header_style), Paragraph("DATE", header_style),
                Paragraph("PARTY NAME", header_style), Paragraph("DESIGN NO", header_style),
                Paragraph("SHIRT COLOUR", header_style), Paragraph("DLY. DATE", header_style),
                Paragraph("PICTURES", header_style)
            ]]
            
            for ord_info in data["orders"]:
                ord_raw = str(ord_info["order_no"]).strip().lower()
                ord_parts = [p for p in re.split(r'[^a-zA-Z0-9]', ord_raw) if p]
                base_order = ord_parts[0] if len(ord_parts) > 0 else ""
                sub_order = ord_parts[1] if len(ord_parts) > 1 else ""
                
                matched_files_data = [] 
                if os.path.exists(ORDER_FOLDER):
                    for f in os.listdir(ORDER_FOLDER):
                        if f.lower().endswith(('.png', '.jpg', '.jpeg', '.jfif', '.bmp', '.webp')):
                            file_name_raw = os.path.splitext(f)[0].lower()
                            file_parts = [p for p in re.split(r'[^a-zA-Z0-9]', file_name_raw) if p]
                            
                            file_base = file_parts[0] if len(file_parts) > 0 else ""
                            file_sub = file_parts[1] if len(file_parts) > 1 else ""
                            
                            matched = False
                            if base_order and file_base == base_order:
                                if sub_order:
                                    if file_sub == sub_order:
                                        matched = True
                                else:
                                    matched = True

                            if matched:
                                img_full_path = os.path.join(ORDER_FOLDER, f)
                                dot_count = file_name_raw.count('.')
                                matched_files_data.append((dot_count, len(file_name_raw), file_name_raw, img_full_path))

                matched_files_data = sorted(list(set(matched_files_data)), key=lambda x: (x[0], x[1]))
                images_list = []
                for item in matched_files_data[:2]:
                    img_full_path = item[3]
                    try:
                        im = PILImage.open(img_full_path)
                        if im.mode in ("RGBA", "P"): im = im.convert("RGB")
                        temp_converted_path = os.path.join(ORDER_FOLDER, f"temp_{item[2].replace('.', '_')}.png")
                        im.save(temp_converted_path, "PNG")
                        images_list.append(temp_converted_path)
                    except:
                        images_list.append(img_full_path)

                img_obj = Paragraph("No Image", cell_style)
                if images_list:
                    img_elements = [RLImage(p, width=38, height=38) for p in images_list if os.path.exists(p)]
                    if img_elements:
                        img_obj = Table([img_elements], colWidths=[45] * len(img_elements))
                        img_obj.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))

                table_data.append([
                    Paragraph(str(ord_info["order_no"]), cell_style), Paragraph(str(ord_info["date"]), cell_style),
                    Paragraph(str(ord_info["party"]), cell_style), Paragraph(str(ord_info["design"]), cell_style),
                    Paragraph(str(ord_info["shirt"]), cell_style), Paragraph(str(ord_info["dly_date"]), cell_style),
                    img_obj
                ])

            t = Table(table_data, colWidths=[65, 65, 120, 80, 80, 65, 95])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4e78')),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6), ('TOPPADDING', (0, 0), (-1, -1), 6),
                ('GRID', (0, 0), (-1, -1), 0.75, colors.HexColor('#333333')),
                ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#1f4e78')),
            ]))

            elements.append(t)
            doc.build(elements)
            absolute_pdf_path = os.path.abspath(pdf_path)

            pyautogui.hotkey('ctrl', 'alt', '/')
            time.sleep(1)
            pyperclip.copy(party_name)
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(2.5)
            pyautogui.press('enter')
            time.sleep(2)

            subprocess.run(f'powershell -Command "Set-Clipboard -Path \'{absolute_pdf_path}\'"', shell=True)
            time.sleep(1.5)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(2.5)
            pyautogui.press('enter')
            time.sleep(3)

            pyperclip.copy("status pls")
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
            pyautogui.press('enter')

            success_count += 1
            time.sleep(2)

        messagebox.showinfo("Completed", f"Successfully {success_count} parties ke liye WhatsApp automation complete ho gayi!")

if __name__ == "__main__":
    root = tk.Tk()
    app = LoginWindow(root)
    root.mainloop()