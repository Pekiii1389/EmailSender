# EmailSender

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import csv  # Importing csv module to handle CSV files


class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Sender")
        self.root.geometry("600x600")
        self.root.resizable(False, False)  # Make the window unsizable

        # Center the window on the screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (1000 // 2)
        y = (screen_height // 2) - (800 // 2)
        self.root.geometry(f"1000x750+{x}+{y}")

        # Set the app's title and style
        self.email_label = tk.Label(self.root, text="Email Sender", font=("Arial", 20, "bold"), fg="#333", bg="#f0f0f0")
        self.email_label.pack(pady=20)

        # Create frames
        self.left_frame = tk.Frame(self.root)
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.right_frame = tk.Frame(self.root)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Email details in the left frame
        self.to_label = tk.Label(self.left_frame, text="To (line separated):", font=("Arial", 12), fg="#333", bg="#f0f0f0")
        self.to_label.pack(pady=5)

        # Frame for email input and scrollbar
        self.to_email_frame = tk.Frame(self.left_frame)
        self.to_email_frame.pack(pady=5)

        self.to_email = tk.Text(self.to_email_frame, width=50, height=25, font=("Arial", 12))
        self.to_email.pack(side=tk.LEFT, fill=tk.BOTH)

        self.to_email_scrollbar = tk.Scrollbar(self.to_email_frame, command=self.to_email.yview)
        self.to_email_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.to_email.config(yscrollcommand=self.to_email_scrollbar.set)

        self.example_to_label = tk.Label(self.left_frame, text="Example:\nemail1@example.com\nemail2@example.com\n...", font=("Arial", 10), fg="#888", bg="#f0f0f0")
        self.example_to_label.pack(pady=5)

        self.send_button = tk.Button(self.left_frame, text="Import from CSV", font=("Arial", 14), command=self.import_from_csv, bg="yellow", fg="black", width=40)
        self.send_button.pack(pady=21)


        # Right frame

        self.subject_label = tk.Label(self.right_frame, text="Subject:", font=("Arial", 12), fg="#333", bg="#f0f0f0")
        self.subject_label.pack(pady=5)
        self.subject_entry = tk.Entry(self.right_frame, width=50, font=("Arial", 12))
        self.subject_entry.pack(pady=5)

        self.body_label = tk.Label(self.right_frame, text="Body:", font=("Arial", 12), fg="#333", bg="#f0f0f0")
        self.body_label.pack(pady=5)

        # Frame for body text and scrollbar
        self.body_frame = tk.Frame(self.right_frame)
        self.body_frame.pack(pady=5)

        self.body_entry = tk.Text(self.body_frame, width=50, height=8, font=("Arial", 12))
        self.body_entry.pack(side=tk.LEFT, fill=tk.BOTH)

        self.body_scrollbar = tk.Scrollbar(self.body_frame, command=self.body_entry.yview)
        self.body_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.body_entry.config(yscrollcommand=self.body_scrollbar.set)

        # Frame for listbox and scrollbar
        self.body_label = tk.Label(self.right_frame, text="Attached Files:", font=("Arial", 12), fg="#333", bg="#f0f0f0")
        self.body_label.pack(pady=5)

        self.file_frame = tk.Frame(self.right_frame)
        self.file_frame.pack(pady=5)

        self.file_listbox = tk.Listbox(self.file_frame, width=50, height=10, font=("Arial", 12), selectmode=tk.MULTIPLE)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH)

        self.scrollbar = tk.Scrollbar(self.file_frame)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_listbox.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.file_listbox.yview)

        self.attach_button = tk.Button(self.right_frame, text="Attach Files", font=("Arial", 14), command=self.attach_files, bg="#4CAF50", fg="white", width=40)
        self.attach_button.pack(pady=5)

        self.withdraw_button = tk.Button(self.right_frame, text="Withdraw Selected File(s)", font=("Arial", 14), command=self.withdraw_file, bg="#FF5722", fg="white", width=40)
        self.withdraw_button.pack(pady=5)

        self.send_button = tk.Button(self.right_frame, text="Send Email", font=("Arial", 14), command=self.confirm_send_email, bg="#4CAF50", fg="white", width=40)
        self.send_button.pack(pady=5)

        self.file_paths = []  # List to store multiple file paths

    def import_from_csv(self):
        file_path = filedialog.askopenfilename(title="Select CSV File", filetypes=(("CSV Files", "*.csv"), ("All Files", "*.*")))
        if file_path:
            with open(file_path, newline='') as csvfile:
                reader = csv.reader(csvfile)
                emails = [row[0] for row in reader if row]  # Assuming emails are in the first column
                self.to_email.delete("1.0", tk.END)  # Clear existing emails
                self.to_email.insert(tk.END, "\n".join(emails))  # Insert emails into the textbox

    def attach_files(self):
        file_paths = filedialog.askopenfilenames(title="Attach files", filetypes=(("All Files", "*.*"),))
        if file_paths:
            for file_path in file_paths:
                if file_path not in self.file_paths:  # Check for duplicates
                    self.file_paths.append(file_path)
                    self.file_listbox.insert(tk.END, os.path.basename(file_path))  # Add file name to listbox
        # No message box displayed on cancel
        else:
            pass

    def withdraw_file(self):
        selected_indices = self.file_listbox.curselection()
        if selected_indices:
            if messagebox.askyesno("Confirm Withdrawal", "Are you sure you want to withdraw the selected files?"):
                for index in selected_indices[::-1]:  # Withdraw files in reverse order to avoid index shifting
                    file_name = self.file_listbox.get(index)
                    file_path = self.file_paths[index]
                    self.file_paths.remove(file_path)
                    self.file_listbox.delete(index)  # Remove from listbox
        else:
            messagebox.showerror("Error", "No file selected to withdraw.")

    def confirm_send_email(self):
        if messagebox.askyesno("Confirm Send", "Are you sure you want to send this email?"):
            self.send_email()

    def send_email(self):
        email_address = "petarmrsa@gmail.com"
        email_password = "clpd imqw eikh vqit"

        recipient_emails = self.to_email.get("1.0", tk.END).strip().split("\n")  # Get emails from Text widget
        subject = self.subject_entry.get()
        body = self.body_entry.get("1.0", tk.END).strip().replace("\n", "<br>")  # Replace newlines with <br> for HTML
        
        if not recipient_emails or not subject or not body:
            messagebox.showerror("Error", "Please fill in all the fields.")
            return

        # Prepare email
        msg = MIMEMultipart()
        msg["From"] = email_address
        msg["To"] = ", ".join(recipient_emails)  # Join multiple recipients
        msg["Subject"] = subject

        # Create HTML body
        html_body = f"""
        <html>
        <head>
            <title>{subject}</title>
        </head>
        <body>
            <p>{body}</p>
        </body>
        </html>
        """
        msg.attach(MIMEText(html_body, "html"))

        # Attach files if selected
        for file_path in self.file_paths:
            if os.path.exists(file_path):
                with open(file_path, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(file_path)}"')
                    msg.attach(part)
            else:
                messagebox.showerror("Error", f"File not found: {file_path}")
                return

        try:
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(email_address, email_password)
            server.sendmail(email_address, recipient_emails, msg.as_string())
            server.quit()
            messagebox.showinfo("Success", "Email sent successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Error: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()
