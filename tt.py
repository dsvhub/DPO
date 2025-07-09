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

DB_FILE = "database.db"
FILE_DIR = "files"

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
        win.geometry("500x400")
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

################################
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
                EMAIL_ADDRESS = "guessacct718@gmail.com"
                EMAIL_PASSWORD = "your_app_password"

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

                # Save to client table if it's new
                name_guess = recipient_var.get().split(" <")[0] if "<" in recipient_var.get() else ""
                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                c.execute("INSERT OR IGNORE INTO clients (email, name, date_added) VALUES (?, ?, ?)",
                          (to_email, name_guess, datetime.now().strftime("%Y-%m-%d")))
                conn.commit()
                conn.close()

                messagebox.showinfo("Success", f"Email sent to {to_email}")
                win.destroy()

            except Exception as e:
                messagebox.showerror("Error", f"Failed to send email:\n{e}")

        tb.Button(win, text="Send Email", bootstyle=SUCCESS, command=send_action).pack(pady=10)

    def view_clients(self):
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT id, name, email, date_added FROM clients ORDER BY date_added DESC")
        clients = c.fetchall()
        conn.close()

        win = tb.Toplevel(self.root)
        win.title("Saved Clients")
        win.geometry("700x300")
        win.grab_set()

        columns = ("Name", "Email", "Date Added")
        tree = tb.Treeview(win, columns=columns, show="headings", height=6, bootstyle="info")
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=220, anchor="w")
        tree.pack(fill="x", expand=True, padx=10, pady=10)

        for cid, name, email, date_added in clients:
            tree.insert("", "end", iid=cid, values=(name or "(No Name)", email, date_added))

        # Right-click menu
        menu = tb.Menu(win, tearoff=0)
        menu.add_command(label="Edit", command=lambda: self.edit_client(tree))
        menu.add_command(label="Delete", command=lambda: self.delete_client(tree))

        def show_menu(event):
            if tree.identify_row(event.y):
                tree.selection_set(tree.identify_row(event.y))
                menu.tk_popup(event.x_root, event.y_root)

        tree.bind("<Button-3>", show_menu)

        # Double-click to reuse email with templates
        def reuse_email(event):
            selected = tree.selection()
            if not selected:
                return
            item = tree.item(selected[0])
            name, email = item['values'][0], item['values'][1]

            # Must select a product from the main table
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

            # Send dialog
            dialog = tb.Toplevel(self.root)
            dialog.title(f"Send to {email}")
            dialog.geometry("550x420")
            dialog.grab_set()

            tb.Label(dialog, text="Subject:").pack(anchor="w", padx=10, pady=(10, 0))
            subject_var = tb.StringVar(value=f"Sharing: {title}")
            tb.Entry(dialog, textvariable=subject_var, width=70).pack(padx=10, pady=5)

            # Template dropdown
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("SELECT title, body FROM templates ORDER BY date_added DESC")
            templates = c.fetchall()
            conn.close()

            tb.Label(dialog, text="Use Template:").pack(anchor="w", padx=10, pady=(5, 0))
            template_var = tb.StringVar()
            template_dropdown = tb.Combobox(dialog, textvariable=template_var,
                                            values=[t[0] for t in templates], width=68)
            template_dropdown.pack(padx=10, pady=5)

            import tkinter as tk
            tb.Label(dialog, text="Message:").pack(anchor="w", padx=10)
            message_box = tk.Text(dialog, height=10, wrap="word")
            message_box.insert("1.0", "Please find the attached file.")
            message_box.pack(padx=10, pady=5, fill="both", expand=True)

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
                body = message_box.get("1.0", "end").strip()

                if not subject or not body:
                    messagebox.showwarning("Missing", "Subject and message are required.")
                    return

                try:
                    EMAIL_ADDRESS = "your.email@gmail.com"
                    EMAIL_PASSWORD = "your_app_password"

                    msg = EmailMessage()
                    msg['Subject'] = subject
                    msg['From'] = EMAIL_ADDRESS
                    msg['To'] = email
                    msg.set_content(body)

                    with open(filepath, 'rb') as f:
                        file_data = f.read()
                        file_name = os.path.basename(filepath)
                        msg.add_attachment(file_data, maintype='application', subtype='octet-stream',
                                           filename=file_name)

                    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                        smtp.starttls()
                        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                        smtp.send_message(msg)

                    messagebox.showinfo("Success", f"Email sent to {email}")
                    dialog.destroy()

                except Exception as e:
                    messagebox.showerror("Error", f"Failed to send email:\n{e}")

            tb.Button(dialog, text="Send Email", bootstyle=SUCCESS, command=send_email).pack(pady=10)

        tree.bind("<Double-1>", reuse_email)

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

        # Reuse send logic
        def send_selected_product(to_email, subject, message_body):
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

            try:
                title, filepath = rows[index]
                EMAIL_ADDRESS = "your.email@gmail.com"
                EMAIL_PASSWORD = "your_app_password"

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

                messagebox.showinfo("Success", f"Email sent to {to_email}")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to send email:\n{e}")

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


# Run app
if __name__ == "__main__":
    init_db()
    root = tb.Window(themename="flatly")
    app = ProductOrganizerApp(root)
    root.mainloop()
