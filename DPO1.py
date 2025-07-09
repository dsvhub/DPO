import os
import sqlite3
import shutil
import subprocess
import csv
from datetime import datetime
import smtplib
from email.message import EmailMessage

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from tkinter import filedialog, messagebox, simpledialog

from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas


import tkinter as tk

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader

from dotenv import load_dotenv

load_dotenv()  # Load environment variables from .env
EMAIL_ADDRESS = os.getenv("APP_EMAIL")
EMAIL_PASSWORD = os.getenv("APP_PASSWORD")


DB_FILE = "database.db"
FILE_DIR = "files"

########################
RECEIPT_DIR = os.path.join(FILE_DIR, "receipts")
os.makedirs(RECEIPT_DIR, exist_ok=True)
########################
CLIENT_FILES_DIR = os.path.join(FILE_DIR, "ClientFiles")
os.makedirs(CLIENT_FILES_DIR, exist_ok=True)

os.makedirs(FILE_DIR, exist_ok=True)

# Initialize DB
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    # Product table
    c.execute('''CREATE TABLE IF NOT EXISTS products (
                    id INTEGER PRIMARY KEY,
                    title TEXT,
                    tags TEXT,
                    category TEXT,
                    filepath TEXT,
                    date_added TEXT
                )''')
    # Client table
    c.execute('''CREATE TABLE IF NOT EXISTS clients (
                    id INTEGER PRIMARY KEY,
                    email TEXT UNIQUE,
                    name TEXT,
                    date_added TEXT
                )''')
    # Email templates table
    c.execute('''CREATE TABLE IF NOT EXISTS templates (
                    id INTEGER PRIMARY KEY,
                    title TEXT,
                    body TEXT,
                    date_added TEXT
                )''')

    conn.commit()
    conn.close()


# Main App
class ProductOrganizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Digital Product Organizer")
        self.root.geometry("850x600")
        self.style = tb.Style("flatly")  # You can change theme

        # Search bar
        self.search_var = tb.StringVar()
        tb.Entry(root, textvariable=self.search_var).pack(fill=X, padx=10, pady=5)
        tb.Button(root, text="Search", bootstyle=INFO, command=self.search_products).pack(pady=(0, 10))

        # Treeview table
        columns = ("Title", "Tags", "Category", "File")
        self.tree = tb.Treeview(root, columns=columns, show="headings", height=15, bootstyle="info")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=180, anchor="w")
        self.tree.pack(padx=10, pady=5)
        self.tree.bind("<Double-1>", self.open_selected_file)

        # Buttons
        button_frame = tb.Frame(root)
        button_frame.pack(pady=10)

        tb.Button(button_frame, text="Add Product", bootstyle=PRIMARY, command=self.add_product).pack(side=LEFT, padx=5)
        tb.Button(button_frame, text="Export CSV", bootstyle=SECONDARY, command=self.export_csv).pack(side=LEFT, padx=5)
        tb.Button(button_frame, text="Send via Email", bootstyle=SUCCESS, command=self.send_email).pack(side=LEFT, padx=5)
        tb.Button(button_frame, text="View Clients", bootstyle=INFO, command=self.view_clients).pack(side=LEFT, padx=5)
        tb.Button(button_frame, text="Delete Product", bootstyle=DANGER, command=self.delete_product).pack(side=LEFT,padx=5)
        tb.Button(button_frame, text="Email Templates", bootstyle=WARNING, command=self.manage_templates).pack(
            side=LEFT, padx=5)

        self.refresh_products()

    def add_product(self):
        filepath = filedialog.askopenfilename()
        if not filepath:
            return

        title = simpledialog.askstring("Title", "Enter product title:")
        tags = simpledialog.askstring("Tags", "Enter tags (comma-separated):")
        category = simpledialog.askstring("Category", "Enter category:")

        if not title:
            messagebox.showerror("Error", "Title is required.")
            return

        filename = os.path.basename(filepath)
        dest_path = os.path.join(FILE_DIR, filename)
        shutil.copy(filepath, dest_path)

        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("INSERT INTO products (title, tags, category, filepath, date_added) VALUES (?, ?, ?, ?, ?)",
                  (title, tags, category, dest_path, datetime.now().strftime("%Y-%m-%d")))
        conn.commit()
        conn.close()

        self.refresh_products()

    def delete_product(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No selection", "Please select a product to delete.")
            return

        confirm = messagebox.askyesno("Delete Confirmation",
                                      "Are you sure you want to delete this product and its file?")
        if not confirm:
            return

        index = self.tree.index(selected[0])

        try:
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("SELECT id, filepath FROM products ORDER BY date_added DESC")
            rows = c.fetchall()
            conn.close()

            product_id, filepath = rows[index]

            # Delete from DB
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("DELETE FROM products WHERE id = ?", (product_id,))
            conn.commit()
            conn.close()

            # Delete file
            if os.path.exists(filepath):
                os.remove(filepath)

            messagebox.showinfo("Deleted", "Product and file deleted successfully.")
            self.refresh_products()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete product:\n{e}")

    def refresh_products(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT id, title, tags, category, filepath FROM products ORDER BY date_added DESC")
        self.filepaths = []  # Track actual paths
        for row in c.fetchall():
            title, tags, category, filepath = row[1], row[2], row[3], os.path.basename(row[4])
            self.tree.insert("", tb.END, values=(title, tags, category, filepath))
            self.filepaths.append(row[4])
        conn.close()

    def open_selected_file(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return

        index = self.tree.index(selected[0])
        try:
            filepath = self.filepaths[index]
            subprocess.Popen(['start', filepath], shell=True)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {e}")

    def search_products(self):
        keyword = self.search_var.get().lower()
        self.tree.delete(*self.tree.get_children())
        self.filepaths = []

        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT title, tags, category, filepath FROM products")
        for row in c.fetchall():
            title, tags, category, path = row
            if keyword in title.lower() or keyword in tags.lower() or keyword in category.lower():
                display = (title, tags, category, os.path.basename(path))
                self.tree.insert("", tb.END, values=display)
                self.filepaths.append(path)
        conn.close()

    def export_csv(self):
        export_path = filedialog.asksaveasfilename(defaultextension=".csv")
        if not export_path:
            return
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT title, tags, category, filepath, date_added FROM products")
        rows = c.fetchall()
        conn.close()

        with open(export_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(["Title", "Tags", "Category", "File", "Date Added"])
            writer.writerows(rows)

        messagebox.showinfo("Exported", f"Exported {len(rows)} products to {export_path}")
#############################################
    def send_email(self):
        import smtplib
        from email.message import EmailMessage

        # Ensure a product is selected
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No selection", "Please select a product to send.")
            return

        index = self.tree.index(selected[0])
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT title, filepath FROM products ORDER BY date_added DESC")
        rows = c.fetchall()
        conn.close()

        if index >= len(rows):
            messagebox.showerror("Error", "Selected product not found.")
            return

        title, filepath = rows[index]

        # ---- Create dialog window ----
        win = tb.Toplevel(self.root)
        win.title("Send Product via Email")
        win.geometry("700x800")
        win.grab_set()

        # Get saved clients
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT name, email FROM clients ORDER BY date_added DESC")
        clients = c.fetchall()
        conn.close()

        emails = [f"{name or '(No Name)'} <{email}>" for name, email in clients]
        email_lookup = {f"{name or '(No Name)'} <{email}>": email for name, email in clients}

        # Dropdown + manual email entry
        tb.Label(win, text="Recipient:").pack(anchor="w", padx=10, pady=(10, 0))
        recipient_var = tb.StringVar()
        dropdown = tb.Combobox(win, textvariable=recipient_var, values=emails, width=60)
        dropdown.pack(padx=10, pady=5)

        tb.Label(win, text="Or type email manually:").pack(anchor="w", padx=10, pady=(10, 0))
        manual_var = tb.StringVar()
        tb.Entry(win, textvariable=manual_var, width=60).pack(padx=10, pady=5)

        # Subject
        tb.Label(win, text="Subject:").pack(anchor="w", padx=10, pady=(10, 0))
        subject_var = tb.StringVar(value=f"Sharing: {title}")
        tb.Entry(win, textvariable=subject_var, width=60).pack(padx=10, pady=5)

        # Message
        tb.Label(win, text="Message:").pack(anchor="w", padx=10, pady=(10, 0))
        message_box = tb.Text(win, height=6, wrap="word")
        message_box.insert("1.0", "Please find the attached file.")
        message_box.pack(padx=10, pady=5, fill="both", expand=False)

        # Load templates
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT title, body FROM templates ORDER BY date_added DESC")
        templates = c.fetchall()
        conn.close()

        tb.Label(win, text="Use Template:").pack(anchor="w", padx=10, pady=(10, 0))
        template_var = tb.StringVar()
        template_dropdown = tb.Combobox(win, textvariable=template_var,
                                        values=[title for title, body in templates],
                                        width=60)
        template_dropdown.pack(padx=10, pady=5)

        def fill_template(event=None):
            selected_title = template_var.get()
            for title, body in templates:
                if title == selected_title:
                    message_box.delete("1.0", "end")
                    message_box.insert("1.0", body)
                    break

        template_dropdown.bind("<<ComboboxSelected>>", fill_template)

        # Price, discount, tax inputs
        tb.Label(win, text="Price:").pack(anchor="w", padx=10, pady=(10, 0))
        price_var = tb.StringVar()
        tb.Entry(win, textvariable=price_var, width=20).pack(padx=10, pady=5, anchor="w")

        tb.Label(win, text="Discount:").pack(anchor="w", padx=10, pady=(10, 0))
        discount_var = tb.StringVar(value="0.0")
        tb.Entry(win, textvariable=discount_var, width=20).pack(padx=10, pady=5, anchor="w")

        tb.Label(win, text="Tax (%):").pack(anchor="w", padx=10, pady=(10, 0))
        tax_var = tb.StringVar(value="0.0")
        tb.Entry(win, textvariable=tax_var, width=20).pack(padx=10, pady=5, anchor="w")

        def send_action():
            selected_email = email_lookup.get(recipient_var.get())
            manual_email = manual_var.get().strip()
            to_email = manual_email or selected_email

            if not to_email:
                messagebox.showerror("Missing", "Please select or type a recipient email.")
                return

            subject = subject_var.get()
            message_body = message_box.get("1.0", "end").strip()

            try:
                EMAIL_ADDRESS = os.getenv("APP_EMAIL")
                EMAIL_PASSWORD = os.getenv("APP_PASSWORD")

                msg = EmailMessage()
                msg['Subject'] = subject
                msg['From'] = EMAIL_ADDRESS
                msg['To'] = to_email
                msg.set_content(message_body)

                with open(filepath, 'rb') as f:
                    file_data = f.read()
                    file_name = os.path.basename(filepath)

                msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

                with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                    smtp.starttls()
                    smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                    smtp.send_message(msg)

                name_guess = recipient_var.get().split(" <")[0] if "<" in recipient_var.get() else ""
                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                c.execute("INSERT OR IGNORE INTO clients (email, name, date_added) VALUES (?, ?, ?)",
                          (to_email, name_guess, datetime.now().strftime("%Y-%m-%d")))
                conn.commit()
                conn.close()

                if price_var.get():
                    try:
                        price = float(price_var.get())
                        discount = float(discount_var.get()) if discount_var.get() else 0.0
                        tax_rate = float(tax_var.get()) / 100 if tax_var.get() else 0.0
                        client_name = name_guess or to_email.split("@")[0]

                        # Ask user to select a logo image (optional)
                        logo_path = filedialog.askopenfilename(
                            title="Select Logo Image (Optional)",
                            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.gif"), ("All files", "*.*")]
                        )
                        if not logo_path:
                            logo_path = None

                        receipt_path = self.generate_receipt(
                            client_name=client_name,
                            client_email=to_email,
                            items=[(file_name, price)],
                            discount=discount,
                            tax=tax_rate,
                            logo_path=logo_path
                        )

                        try:
                            subprocess.Popen(['start', receipt_path], shell=True)
                        except:
                            pass

                    except Exception as e:
                        print("Failed to generate receipt:", e)

                messagebox.showinfo("Success", f"Email sent to {to_email}\nReceipt saved.")
                win.destroy()

            except Exception as e:
                messagebox.showerror("Error", f"Failed to send email:\n{e}")

        tb.Button(win, text="Send Email", bootstyle=SUCCESS, command=send_action).pack(pady=10)

##################################################
    def manage_client_files(self, tree):
        selected = tree.selection()
        if not selected:
            return

        item = tree.item(selected[0])
        client_name = item['values'][0]
        email = item['values'][1]
        client_folder = os.path.join(CLIENT_FILES_DIR, client_name.replace(" ", "_"))
        os.makedirs(client_folder, exist_ok=True)

        win = tb.Toplevel(self.root)
        win.title(f"Files for {client_name}")
        win.geometry("650x450")
        win.grab_set()

        file_list = tk.Listbox(win)
        file_list.pack(fill="both", expand=True, padx=10, pady=10)

        def refresh():
            file_list.delete(0, "end")
            for f in os.listdir(client_folder):
                file_list.insert("end", f)

        def add_file():
            path = filedialog.askopenfilename()
            if path:
                shutil.copy(path, os.path.join(client_folder, os.path.basename(path)))
                refresh()

        def delete_file():
            sel = file_list.curselection()
            if not sel:
                return
            filename = file_list.get(sel[0])
            confirm = messagebox.askyesno("Confirm", f"Delete file '{filename}'?")
            if confirm:
                os.remove(os.path.join(client_folder, filename))
                refresh()

        def send_file():
            sel = file_list.curselection()
            if not sel:
                return
            filename = file_list.get(sel[0])
            filepath = os.path.join(client_folder, filename)

            # Create email window
            send_win = tb.Toplevel(win)
            send_win.title("Send File")
            send_win.geometry("500x600")
            send_win.grab_set()

            # Subject
            tb.Label(send_win, text="Subject:").pack(anchor="w", padx=10, pady=(10, 0))
            subject_var = tb.StringVar(value=f"Sharing: {filename}")
            tb.Entry(send_win, textvariable=subject_var, width=60).pack(padx=10, pady=5)

            # Message body
            tb.Label(send_win, text="Message:").pack(anchor="w", padx=10)
            message_box = tb.Text(send_win, height=6, wrap="word")
            message_box.insert("1.0", "Please find the attached file.")
            message_box.pack(padx=10, pady=5, fill="both", expand=False)

            # Templates
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("SELECT title, body FROM templates ORDER BY date_added DESC")
            templates = c.fetchall()
            conn.close()

            tb.Label(send_win, text="Use Template:").pack(anchor="w", padx=10, pady=(10, 0))
            template_var = tb.StringVar()
            template_dropdown = tb.Combobox(send_win, textvariable=template_var,
                                            values=[title for title, _ in templates], width=60)
            template_dropdown.pack(padx=10, pady=5)

            def fill_template(event=None):
                selected_title = template_var.get()
                for title, body in templates:
                    if title == selected_title:
                        message_box.delete("1.0", "end")
                        message_box.insert("1.0", body)
                        break

            template_dropdown.bind("<<ComboboxSelected>>", fill_template)

            # Price, discount, tax inputs
            tb.Label(send_win, text="Price:").pack(anchor="w", padx=10, pady=(10, 0))
            price_var = tb.StringVar()
            tb.Entry(send_win, textvariable=price_var, width=20).pack(padx=10, pady=5, anchor="w")

            tb.Label(send_win, text="Discount:").pack(anchor="w", padx=10, pady=(10, 0))
            discount_var = tb.StringVar(value="0.0")
            tb.Entry(send_win, textvariable=discount_var, width=20).pack(padx=10, pady=5, anchor="w")

            tb.Label(send_win, text="Tax (%):").pack(anchor="w", padx=10, pady=(10, 0))
            tax_var = tb.StringVar(value="0.0")
            tb.Entry(send_win, textvariable=tax_var, width=20).pack(padx=10, pady=5, anchor="w")

            # Logo picker
            tb.Label(send_win, text="Select Logo (optional):").pack(anchor="w", padx=10, pady=(10, 0))
            logo_path_var = tk.StringVar()
            tb.Entry(send_win, textvariable=logo_path_var, width=60).pack(padx=10, pady=5, anchor="w")

            def browse_logo():
                file = filedialog.askopenfilename(
                    title="Select Logo Image",
                    filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.gif"), ("All files", "*.*")]
                )
                if file:
                    logo_path_var.set(file)

            tb.Button(send_win, text="Browse Logo", command=browse_logo).pack(padx=10, pady=(0, 10), anchor="w")

            def send_action():
                subject = subject_var.get().strip()
                message_body = message_box.get("1.0", "end").strip()

                if not subject or not message_body:
                    messagebox.showwarning("Missing", "Subject and message are required.")
                    return

                try:
                    EMAIL_ADDRESS = os.getenv("APP_EMAIL")
                    EMAIL_PASSWORD = os.getenv("APP_PASSWORD")

                    msg = EmailMessage()
                    msg['Subject'] = subject
                    msg['From'] = EMAIL_ADDRESS
                    msg['To'] = email
                    msg.set_content(message_body)

                    with open(filepath, 'rb') as f:
                        file_data = f.read()
                        msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=filename)

                    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                        smtp.starttls()
                        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                        smtp.send_message(msg)

                    # Generate receipt if price was entered
                    if price_var.get():
                        try:
                            price = float(price_var.get())
                            discount = float(discount_var.get()) if discount_var.get() else 0.0
                            tax_rate = float(tax_var.get()) if tax_var.get() else 0.0

                            items = [(filename, price)]
                            receipt_path = self.generate_receipt(
                                client_name=client_name,
                                client_email=email,
                                items=items,
                                discount=discount,
                                tax=tax_rate,
                                logo_path=logo_path_var.get() or None
                            )
                            subprocess.Popen(['start', receipt_path], shell=True)

                        except Exception as e:
                            print("Failed to generate receipt PDF:", e)

                    messagebox.showinfo("Success", f"Email sent to {email}")
                    send_win.destroy()

                except Exception as e:
                    messagebox.showerror("Error", f"Failed to send email:\n{e}")

            tb.Button(send_win, text="Send Email", bootstyle=SUCCESS, command=send_action).pack(pady=10)

        # Action buttons
        btn_frame = tb.Frame(win)
        btn_frame.pack(pady=5)

        tb.Button(btn_frame, text="Add File", bootstyle=PRIMARY, command=add_file).pack(side="left", padx=10)
        tb.Button(btn_frame, text="Delete File", bootstyle=DANGER, command=delete_file).pack(side="left", padx=10)
        tb.Button(btn_frame, text="Send via Email", bootstyle=SUCCESS, command=send_file).pack(side="left", padx=10)

        refresh()

    ################################################

    def view_clients(self):
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT id, name, email, date_added FROM clients ORDER BY date_added DESC")
        clients = c.fetchall()
        conn.close()

        win = tb.Toplevel(self.root)
        win.title("Saved Clients")
        win.geometry("600x400")

        columns = ("Name", "Email", "Date Added")
        tree = tb.Treeview(win, columns=columns, show="headings", bootstyle="info")
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=180, anchor="w")
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        # Populate table
        for client in clients:
            cid, name, email, date_added = client
            tree.insert("", "end", iid=cid, values=(name or "(No Name)", email, date_added))

        # Right-click menu
        menu = tb.Menu(win, tearoff=0)
        menu.add_command(label="Edit", command=lambda: edit_client(tree))
        menu.add_command(label="Delete", command=lambda: delete_client(tree))
        menu.add_command(label="Manage Files", command=lambda: self.manage_client_files(tree))

        def show_menu(event):
            if tree.identify_row(event.y):
                tree.selection_set(tree.identify_row(event.y))
                menu.tk_popup(event.x_root, event.y_root)

        tree.bind("<Button-3>", show_menu)
#########################################################################
        # Double-click to reuse email
        def reuse_email(event):
            selected = tree.selection()
            if not selected:
                return
            item = tree.item(selected[0])
            name, email = item['values'][0], item['values'][1]

            main_selected = self.tree.selection()
            if not main_selected:
                messagebox.showinfo("Select Product", "Please select a product from the main list to send.")
                return

            product_index = self.tree.index(main_selected[0])
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("SELECT title, filepath FROM products ORDER BY date_added DESC")
            rows = c.fetchall()
            conn.close()

            if product_index >= len(rows):
                messagebox.showerror("Error", "Product index out of range.")
                return

            title, filepath = rows[product_index]
            file_name = os.path.basename(filepath)

            win = tb.Toplevel(self.root)
            win.title(f"Send to {email}")
            win.geometry("700x700")
            win.grab_set()

            tb.Label(win, text="Subject:").pack(padx=10, pady=(10, 0), anchor="w")
            subject_var = tb.StringVar(value=f"Sharing: {title}")
            tb.Entry(win, textvariable=subject_var, width=60).pack(padx=10, pady=5)

            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("SELECT title, body FROM templates ORDER BY date_added DESC")
            templates = c.fetchall()
            conn.close()

            tb.Label(win, text="Use Template:").pack(anchor="w", padx=10, pady=(10, 0))
            template_var = tb.StringVar()
            template_dropdown = tb.Combobox(win, textvariable=template_var,
                                            values=[t[0] for t in templates], width=60)
            template_dropdown.pack(padx=10, pady=5)

            tb.Label(win, text="Message:").pack(anchor="w", padx=10)
            message_box = tk.Text(win, height=6, wrap="word")
            message_box.insert("1.0", "Please find the attached file.")
            message_box.pack(padx=10, pady=5, fill="both", expand=True)

            tb.Label(win, text="Price:").pack(padx=10, anchor="w")
            price_var = tb.StringVar()
            tb.Entry(win, textvariable=price_var).pack(padx=10, pady=5)

            tb.Label(win, text="Discount:").pack(padx=10, anchor="w")
            discount_var = tb.StringVar(value="0")
            tb.Entry(win, textvariable=discount_var).pack(padx=10, pady=5)

            tb.Label(win, text="Tax (%):").pack(padx=10, anchor="w")
            tax_var = tb.StringVar(value="0")
            tb.Entry(win, textvariable=tax_var).pack(padx=10, pady=5)

            # Logo picker
            logo_path_var = tk.StringVar()

            def pick_logo():
                path = filedialog.askopenfilename(
                    title="Select Logo (optional)",
                    filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.gif"), ("All files", "*.*")]
                )
                if path:
                    logo_path_var.set(path)

            tb.Button(win, text="Select Logo", command=pick_logo).pack(pady=5)
            tb.Label(win, textvariable=logo_path_var, wraplength=500, justify="left").pack(padx=10)

            def fill_template(event=None):
                selected_title = template_var.get()
                for t_title, t_body in templates:
                    if t_title == selected_title:
                        message_box.delete("1.0", "end")
                        message_box.insert("1.0", t_body)
                        break

            template_dropdown.bind("<<ComboboxSelected>>", fill_template)

            def send_email():
                subject = subject_var.get().strip()
                message_body = message_box.get("1.0", "end").strip()

                if not subject or not message_body:
                    messagebox.showwarning("Missing", "Subject and message are required.")
                    return

                try:
                    EMAIL_ADDRESS = os.getenv("APP_EMAIL")
                    EMAIL_PASSWORD = os.getenv("APP_PASSWORD")

                    msg = EmailMessage()
                    msg['Subject'] = subject
                    msg['From'] = EMAIL_ADDRESS
                    msg['To'] = email
                    msg.set_content(message_body)

                    with open(filepath, 'rb') as f:
                        file_data = f.read()

                    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

                    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                        smtp.starttls()
                        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                        smtp.send_message(msg)

                    client_name_final = name or email.split("@")[0]
                    try:
                        price = float(price_var.get()) if price_var.get() else 0.0
                        discount = float(discount_var.get()) if discount_var.get() else 0.0
                        tax = float(tax_var.get()) / 100 if tax_var.get() else 0.0
                        logo_path = logo_path_var.get() if logo_path_var.get() else None

                        if price > 0:
                            items = [(file_name, price)]
                            receipt_path = self.generate_receipt(
                                client_name=client_name_final,
                                client_email=email,
                                items=items,
                                discount=discount,
                                tax=tax,
                                logo_path=logo_path
                            )
                            subprocess.Popen(['start', receipt_path], shell=True)

                    except Exception as e:
                        print("Failed to generate receipt:", e)

                    messagebox.showinfo("Success", f"Email sent to {email}")
                    win.destroy()

                except Exception as e:
                    messagebox.showerror("Error", f"Failed to send email:\n{e}")

            tb.Button(win, text="Send Email", bootstyle=SUCCESS, command=send_email).pack(pady=10)

        tree.bind("<Double-1>", reuse_email)

        ########################################################################

        # Edit client
        def edit_client(treeview):
            selected = treeview.selection()
            if not selected:
                return
            client_id = selected[0]
            current = treeview.item(client_id)["values"]
            new_name = simpledialog.askstring("Edit Name", "Enter new name:", initialvalue=current[0])
            if new_name is None:
                return
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("UPDATE clients SET name = ? WHERE id = ?", (new_name, client_id))
            conn.commit()
            conn.close()
            treeview.item(client_id, values=(new_name, current[1], current[2]))

        # Delete client
        def delete_client(treeview):
            selected = treeview.selection()
            if not selected:
                return
            client_id = selected[0]
            confirm = messagebox.askyesno("Delete", "Are you sure you want to delete this client?")
            if not confirm:
                return
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("DELETE FROM clients WHERE id = ?", (client_id,))
            conn.commit()
            conn.close()
            treeview.delete(client_id)

#######################################################################

        def send_selected_product(to_email=None, subject=None, message_body=None):
            selected = self.tree.selection()
            if not selected:
                messagebox.showwarning("No selection", "Please select a product first in the main list.")
                return

            index = self.tree.index(selected[0])
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("SELECT title, filepath FROM products ORDER BY date_added DESC")
            rows = c.fetchall()
            conn.close()

            if index >= len(rows):
                messagebox.showerror("Error", "Selected product not found.")
                return

            title, filepath = rows[index]
            file_name = os.path.basename(filepath)

            win = tb.Toplevel(self.root)
            win.title("Send Product via Email")
            win.geometry("600x650")
            win.grab_set()

            # Recipient
            tb.Label(win, text="To Email:").pack(padx=10, anchor="w")
            to_var = tk.StringVar(value=to_email or "")
            tb.Entry(win, textvariable=to_var, width=60).pack(padx=10, pady=5)

            # Subject
            tb.Label(win, text="Subject:").pack(padx=10, anchor="w")
            subject_var = tk.StringVar(value=subject or f"Sharing: {title}")
            tb.Entry(win, textvariable=subject_var, width=60).pack(padx=10, pady=5)

            # Message body
            tb.Label(win, text="Message:").pack(padx=10, anchor="w")
            message_box = tk.Text(win, height=6, wrap="word")
            message_box.insert("1.0", message_body or "Please find the attached file.")
            message_box.pack(padx=10, pady=5, fill="both")

            # Template dropdown
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("SELECT title, body FROM templates ORDER BY date_added DESC")
            templates = c.fetchall()
            conn.close()

            tb.Label(win, text="Use Template:").pack(anchor="w", padx=10)
            template_var = tk.StringVar()
            template_dropdown = tb.Combobox(win, textvariable=template_var,
                                            values=[t[0] for t in templates], width=60)
            template_dropdown.pack(padx=10, pady=5)

            def fill_template(event=None):
                for t_title, t_body in templates:
                    if t_title == template_var.get():
                        message_box.delete("1.0", "end")
                        message_box.insert("1.0", t_body)
                        break

            template_dropdown.bind("<<ComboboxSelected>>", fill_template)

            # Price / Discount / Tax inputs
            tb.Label(win, text="Price:").pack(padx=10, anchor="w")
            price_var = tk.StringVar()
            tb.Entry(win, textvariable=price_var, width=20).pack(padx=10, pady=3)

            tb.Label(win, text="Discount (optional):").pack(padx=10, anchor="w")
            discount_var = tk.StringVar(value="0.0")
            tb.Entry(win, textvariable=discount_var, width=20).pack(padx=10, pady=3)

            tb.Label(win, text="Tax % (optional):").pack(padx=10, anchor="w")
            tax_var = tk.StringVar(value="0.0")
            tb.Entry(win, textvariable=tax_var, width=20).pack(padx=10, pady=3)

            # Logo picker button
            logo_path_var = tk.StringVar()

            def pick_logo():
                path = filedialog.askopenfilename(
                    title="Select Logo (optional)",
                    filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.gif"), ("All files", "*.*")]
                )
                if path:
                    logo_path_var.set(path)

            tb.Button(win, text="Select Logo", command=pick_logo).pack(pady=5)
            tb.Label(win, textvariable=logo_path_var, wraplength=500, justify="left").pack(padx=10)

            def send_action():
                to = to_var.get().strip()
                subject = subject_var.get().strip()
                body = message_box.get("1.0", "end").strip()

                if not to or not subject or not body:
                    messagebox.showerror("Missing Data", "Email, subject, and message are required.")
                    return

                try:
                    EMAIL_ADDRESS = os.getenv("APP_EMAIL")
                    EMAIL_PASSWORD = os.getenv("APP_PASSWORD")

                    msg = EmailMessage()
                    msg['Subject'] = subject
                    msg['From'] = EMAIL_ADDRESS
                    msg['To'] = to
                    msg.set_content(body)

                    with open(filepath, 'rb') as f:
                        file_data = f.read()
                    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

                    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                        smtp.starttls()
                        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                        smtp.send_message(msg)

                    client_name = to.split("@")[0]
                    price_str = price_var.get().strip()

                    if price_str:
                        try:
                            price = float(price_str)
                            discount = float(discount_var.get() or 0.0)
                            tax = float(tax_var.get() or 0.0) / 100.0
                            logo_path = logo_path_var.get() if logo_path_var.get() else None

                            receipt_path = self.generate_receipt(
                                client_name=client_name,
                                client_email=to,
                                items=[(file_name, price)],
                                discount=discount,
                                tax=tax,
                                logo_path=logo_path
                            )
                            subprocess.Popen(['start', receipt_path], shell=True)

                        except Exception as e:
                            print("Error generating receipt:", e)
                    else:
                        # fallback text receipt
                        txt_path = os.path.join(FILE_DIR, "receipts", f"{client_name}_{file_name}.txt")
                        os.makedirs(os.path.dirname(txt_path), exist_ok=True)
                        with open(txt_path, "w") as f:
                            f.write(
                                f"Receipt\nClient: {client_name}\nEmail: {to}\nFile: {file_name}\nDate: {datetime.now()}")
                        subprocess.Popen(['start', txt_path], shell=True)

                    messagebox.showinfo("Success", f"Email sent to {to}")
                    win.destroy()

                except Exception as e:
                    messagebox.showerror("Error", f"Failed to send email:\n{e}")

            tb.Button(win, text="Send Email", bootstyle=SUCCESS, command=send_action).pack(pady=15)

    ##################################################################################
    def manage_templates(self):
        win = tb.Toplevel(self.root)
        win.title("Manage Email Templates")
        win.geometry("600x400")
        win.grab_set()

        columns = ("Title", "Body")
        tree = tb.Treeview(win, columns=columns, show="headings", bootstyle="info")
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=250, anchor="w")
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        def refresh():
            tree.delete(*tree.get_children())
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("SELECT id, title, body FROM templates ORDER BY date_added DESC")
            for row in c.fetchall():
                tree.insert("", "end", iid=row[0], values=(row[1], row[2][:100] + "..."))
            conn.close()

        def add_template():
            t_win = tb.Toplevel(win)
            t_win.title("New Template")
            t_win.geometry("500x500")
            t_win.grab_set()

            # Title
            tb.Label(t_win, text="Title:").pack(padx=10, pady=(10, 0), anchor="w")
            title_var = tb.StringVar()
            tb.Entry(t_win, textvariable=title_var, width=60).pack(padx=10, pady=5)

            # Body label
            tb.Label(t_win, text="Body:").pack(padx=10, pady=(10, 0), anchor="w")

            # Use tk.Text explicitly
            import tkinter as tk
            text_box = tk.Text(t_win, height=10, wrap="word")
            text_box.pack(padx=10, pady=5, fill="both", expand=True)

            # Save button
            def save_template():
                title = title_var.get().strip()
                body = text_box.get("1.0", "end").strip()
                if not title or not body:
                    messagebox.showwarning("Missing", "Both title and body are required.")
                    return
                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                c.execute("INSERT INTO templates (title, body, date_added) VALUES (?, ?, ?)",
                          (title, body, datetime.now().strftime("%Y-%m-%d")))
                conn.commit()
                conn.close()
                t_win.destroy()
                refresh()

            tb.Button(t_win, text="Save Template", bootstyle=SUCCESS, command=save_template).pack(pady=10)


        def delete_template():
            selected = tree.selection()
            if not selected:
                return
            confirm = messagebox.askyesno("Delete", "Delete selected template?")
            if not confirm:
                return
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("DELETE FROM templates WHERE id = ?", (selected[0],))
            conn.commit()
            conn.close()
            refresh()

        btn_frame = tb.Frame(win)
        btn_frame.pack(pady=5)

        tb.Button(btn_frame, text="Add", bootstyle=PRIMARY, command=add_template).pack(side=LEFT, padx=5)
        tb.Button(btn_frame, text="Delete", bootstyle=DANGER, command=delete_template).pack(side=LEFT, padx=5)

        refresh()
####################################

    def generate_receipt(
            self,
            client_name,
            client_email,
            items,  # List of tuples like [("File A.pdf", 10.00)]
            discount: float = 0.0,
            tax: float = 0.0,  # percentage (e.g. 8.5)
            extra_save_paths=[],
            logo_path=None
    ):
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        base_filename = f"{client_name.replace(' ', '_')}_{now.replace(':', '_').replace(' ', '_')}"
        receipt_dir = os.path.join(FILE_DIR, "receipts")
        os.makedirs(receipt_dir, exist_ok=True)
        receipt_path = os.path.join(receipt_dir, f"{base_filename}.pdf")

        c = canvas.Canvas(receipt_path, pagesize=LETTER)
        width, height = LETTER
        y = height - 50

        # Try to draw logo (optional)
        if isinstance(logo_path, str) and os.path.exists(logo_path):
            try:
                logo = ImageReader(logo_path)
                logo_width = 120
                logo_height = 60
                c.drawImage(logo, 50, y - logo_height, width=logo_width, height=logo_height, preserveAspectRatio=True,
                            mask='auto')
                y -= logo_height + 20
            except Exception as e:
                print(f"[Receipt] Failed to load logo: {e}")
                y -= 10
        else:
            y -= 10

        # Title
        c.setFont("Helvetica-Bold", 16)
        c.drawString(50, y, "Digital Product Receipt")
        y -= 40

        # Date & Client Info
        c.setFont("Helvetica", 12)
        c.drawString(50, y, f"Date: {now}")
        y -= 20
        c.drawString(50, y, f"Client: {client_name} ({client_email})")
        y -= 30

        # Line separator
        c.line(50, y, width - 50, y)
        y -= 20

        # Table headers
        c.setFont("Helvetica-Bold", 12)
        c.drawString(50, y, "Item")
        c.drawString(300, y, "Price")
        y -= 20

        # Item rows
        c.setFont("Helvetica", 12)
        subtotal = 0.0
        for name, price in items:
            c.drawString(50, y, str(name))
            c.drawString(300, y, f"${float(price):.2f}")
            subtotal += float(price)
            y -= 20

        # Line below items
        c.line(50, y, width - 50, y)
        y -= 20

        # Subtotal
        c.setFont("Helvetica", 12)
        c.drawString(50, y, "Subtotal:")
        c.drawString(300, y, f"${subtotal:.2f}")
        y -= 20

        # Discount
        if discount > 0:
            c.drawString(50, y, "Discount:")
            c.drawString(300, y, f"-${discount:.2f}")
            y -= 20

        # Tax
        tax_amount = round((subtotal - discount) * (tax / 100), 2) if tax > 0 else 0.0
        if tax_amount > 0:
            c.drawString(50, y, f"Tax ({tax:.1f}%):")
            c.drawString(300, y, f"${tax_amount:.2f}")
            y -= 20

        # Grand total
        total = subtotal - discount + tax_amount
        c.setFont("Helvetica-Bold", 12)
        c.drawString(50, y, "Grand Total:")
        c.drawString(300, y, f"${total:.2f}")
        y -= 40

        # Footer
        c.setFont("Helvetica", 11)
        c.drawString(50, y, "Thank you for your business!")

        c.save()

        # Save a copy to extra paths
        for path in extra_save_paths:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            shutil.copy(receipt_path, path)

        return receipt_path


# Run app
if __name__ == "__main__":
    init_db()
    root = tb.Window(themename="flatly")
    app = ProductOrganizerApp(root)
    root.mainloop()
