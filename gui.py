import tkinter as tk
from tkinter import messagebox
from src.helper import send_mail

class EmailGui(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Send Email App")
        self.geometry("400x300")
        self.make_combobox()
    
    def make_combobox(self):
        tk.Label(self, text="Enter RFQ Number: ").grid(row=0, column=0)
        self.rfq_number = tk.Entry(self, width=20)
        self.rfq_number.grid(row=1, column=0)
        send_mail_button = tk.Button(self, text="Send Email", command=lambda: self.verify_and_send_email(self.rfq_number.get()))
        send_mail_button.grid(row=2, column=0)

    def verify_and_send_email(self, rfq_number):
        if rfq_number:
            send_mail(rfq_number)
            messagebox.showinfo("Success", "Email sent successfully")
            self.rfq_number.delete(0, tk.END)
        else:
            messagebox.showerror("Error", "Please enter RFQ number")

if __name__ == "__main__":
    app = EmailGui()
    app.mainloop()
