import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import threading
from tkinter import ttk
from utils import *
import customtkinter


class MailedItApp(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("MailedIt!")
        self.iconbitmap("../assets/logo.ico")
        self._data = None
        self._template_file_path = None
        self.email_label = customtkinter.CTkLabel(self, text="Enter Email:")
        self.email_label.grid(row=0, column=0, padx=10, pady=5)
        self.email_entry = customtkinter.CTkEntry(self, width=300)
        self.email_entry.grid(row=0, column=1, columnspan=2, padx=10, pady=5)
        self.password_label = customtkinter.CTkLabel(self, text="Enter Password:")
        self.password_label.grid(row=1, column=0, padx=10, pady=5)
        self.password_entry = customtkinter.CTkEntry(self, show="*", width=300)
        self.password_entry.grid(row=1, column=1, columnspan=2, padx=10, pady=5)

        self.html_template_label = customtkinter.CTkLabel(
            self, text="Select HTML Template File:"
        )
        self.html_template_label.grid(row=2, column=0, padx=10, pady=5)
        self.html_template_entry = customtkinter.CTkEntry(self)
        self.html_template_entry.grid(row=2, column=1, padx=10, pady=5)
        self.html_template_button = customtkinter.CTkButton(
            self, text="Browse", command=self.browse_html_template
        )
        self.html_template_button.grid(row=2, column=2, padx=5, pady=5)

        self.excel_data_label = customtkinter.CTkLabel(
            self, text="Select Excel Data File:"
        )
        self.excel_data_label.grid(row=3, column=0, padx=10, pady=5)
        self.excel_data_entry = customtkinter.CTkEntry(self)
        self.excel_data_entry.grid(row=3, column=1, padx=10, pady=5)
        self.excel_data_button = customtkinter.CTkButton(
            self, text="Browse", command=self.browse_excel_data
        )
        self.excel_data_button.grid(row=3, column=2, padx=5, pady=5)

        self.submit_button = customtkinter.CTkButton(
            self, text="Start Mailing", command=self.submit
        )
        self.submit_button.grid(row=4, column=0, padx=10, pady=10)

        self.cancel_button = customtkinter.CTkButton(
            self, text="Cancel", command=self.cancel
        )
        self.cancel_button.grid(row=4, column=1, padx=5, pady=10)

        self.help_button = customtkinter.CTkButton(
            self, text="Help", command=self.show_help_window
        )
        self.help_button.grid(row=4, column=2, padx=5, pady=10)

    def browse_html_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("HTML Files", "*.html")])
        self.html_template_entry.delete(0, tk.END)
        self.html_template_entry.insert(0, file_path)
        self._template_file_path = file_path
        if not check_file_existence_and_readability(file_path):
            self.html_template_entry.delete(0, tk.END)
            self._template_file_path = None
            messagebox.showerror("Error", "HTML File is not readable/invalid!")

    def browse_excel_data(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        self.excel_data_entry.delete(0, tk.END)
        self.excel_data_entry.insert(0, file_path)
        self._data = read_excel_to_dict(file_path)
        if self._data is None or len(self._data) == 0:
            self.excel_data_entry.delete(0, tk.END)
            messagebox.showerror("Error", "Invalid excel file!")

    def send_emails(self):
        if self._data is None:
            messagebox.showerror("Error", "Please select an excel data file!")
            return
        if self._template_file_path is None:
            messagebox.showerror("Error", "Please select an HTML template file!")
            return

        def send_emails_process():
            num_emails = len(self._data)
            mailed_emails = []
            password = self.password_entry.get()
            from_address = self.email_entry.get()

            html_template_path = self.html_template_entry.get()

            for i, record in enumerate(self._data):
                to_address = record.get("email")
                subject = record.get("subject")
                message_html = render_html_template(html_template_path, record)
                if message_html is None or to_address is None or subject is None:
                    messagebox.showerror("Error", "Insufficient Data!")
                    return
                try:
                    send_email(
                        from_address, password, to_address, subject, message_html
                    )
                except SMTPAuthenticationError:
                    messagebox.showerror("Error", "Invalid credentials!")
                    return
                except SMTPException:
                    messagebox.showerror("Error", "Error while sending email!")
                    return

                mailed_emails.append(to_address)
                update_table(mailed_emails)
                status_label.config(text=f"Sending email {i + 1}/{num_emails}")

            completion_label = tk.Label(mail_status_window, text="Mailing completed!")
            completion_label.pack(pady=10)

            close_button = customtkinter.CTkButton(
                mail_status_window, text="Close", command=close_mail_status_window
            )
            close_button.pack(pady=5)

        def update_table(emails):
            # Clear the existing table content
            for row in email_table.get_children():
                email_table.delete(row)

            # Insert the mailed emails into the table
            for i, email in enumerate(emails):
                email_table.insert("", "end", values=(i + 1, email))

        # Create a new window to show mailing progress
        mail_status_window = tk.Toplevel(self)
        mail_status_window.title("Mailing Progress")

        def close_mail_status_window():
            mail_status_window.destroy()

        frame_table = ttk.Frame(mail_status_window)
        frame_table.pack(pady=5)
        email_table = ttk.Treeview(
            frame_table,
            columns=("Serial No.", "Email Address"),
            show="headings",
        )
        email_table.heading("Serial No.", text="Serial No.")
        email_table.heading("Email Address", text="Email Address")
        email_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        table_scrollbar = ttk.Scrollbar(
            frame_table, orient=tk.VERTICAL, command=email_table.yview
        )
        table_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        email_table.configure(yscrollcommand=table_scrollbar.set)

        status_label = tk.Label(mail_status_window, text="Sending email...")
        status_label.pack(pady=5)

        threading.Thread(target=send_emails_process).start()

    def submit(self):
        if not self.email_entry.get().endswith("@gmail.com"):
            messagebox.showwarning(
                "Error", "Please enter a valid email ending in gmail.com!"
            )
            return
        if self.password_entry.get() == "":
            messagebox.showwarning("Alert", "Please enter a valid password!")
            return
        threading.Thread(target=self.send_emails).start()

    def cancel(self):
        self.destroy()

    def show_help_window(self):
        help_text = (
            "Help\n\n"
            "1. Obtaining app-specific password from Google account settings in 2FA auth:\n"
            "   - Go to your Google account settings.\n"
            "   - Click on 'Security' in the left navigation pane.\n"
            "   - Under 'Signing in to Google', click on 'App passwords'.\n"
            "   - Select the app you want to generate an app-specific password for.\n"
            "   - Follow the prompts to create the app-specific password.\n\n"
            "2. Template syntax basics of Jinja2 for HTML template:\n"
            "   - Jinja2 is a popular templating language for Python.\n"
            "   - To use Jinja2 in your HTML template, enclose variables within double curly braces {{ }}.\n"
            "   - For example: <p>Hello, {{ name }}!</p>\n"
            "   - The {{ name }} will be replaced with the value of the 'name' variable during rendering.\n\n"
            "3. Structure of Excel file to include each variable created in the template:\n"
            "   - Your Excel file should have a header row with column names corresponding to the variables used in the template.\n"
            "   - Each row represents a data entry, and the values in each cell will be used to replace the corresponding variables in the template.\n"
            "   - For example:\n"
            "       |\tname\t|\tage\t|\n"
            "       |\tJohn\t|\t25\t|\n"
            "       |\tAlice\t|\t30\t|\n"
        )

        help_window = tk.Toplevel(self)
        help_window.title("Help")

        help_text_widget = tk.Text(
            help_window,
            wrap="word",
            width=80,
            height=20,
            padx=10,
            pady=10,
            font=("Helvetica", 12),
        )
        help_text_widget.insert(tk.END, help_text)
        help_text_widget.tag_configure("heading", font=("Helvetica", 14, "bold"))
        help_text_widget.tag_configure("align_left", justify="left")

        help_text_widget.tag_add("heading", "1.0", "1.end")
        help_text_widget.tag_add("align_left", "1.0", "end")

        help_text_widget.config(state=tk.DISABLED)  # Make the Text widget read-only
        help_text_widget.pack(fill=tk.BOTH, expand=True)


if __name__ == "__main__":
    app = MailedItApp()
    app.mainloop()
