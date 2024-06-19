import tkinter as tk
from tkinter import messagebox, filedialog
from src.helper import send_mail

class EmailGui(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Send Email App")
        self.geometry("550x300")
        self.make_combobox()
    
    def make_combobox(self):
        tk.Label(self, text="Enter RFQ Number: ").grid(row=0, column=0)
        self.rfq_number = tk.Entry(self, width=20)
        self.rfq_number.grid(row=1, column=0)

        tk.Label(self, text="Enter Item ID: ").grid(row=0, column=1)
        self.item_id = tk.Entry(self, width=20)
        self.item_id.grid(row=1, column=1)

        tk.Label(self, text="Enter Item Qty Req ").grid(row=2, column=1)
        self.item_qty = tk.Entry(self, width=20)
        self.item_qty.grid(row=3, column=1)
        

        tk.Label(self, text="Other Attachments:").grid(row=2, column=0)
        self.other_attachments = tk.Listbox(self, height=2, width=50)
        self.other_attachments.grid(row=3, column=0)

        browse_button_part_list = tk.Button(self, text="Browse Files", command=lambda: self.browse_files_parts_requested("All files", self.other_attachments))
        browse_button_part_list.grid(row=4, column=0)

        send_mail_button = tk.Button(self, text="Send Email", command=lambda: self.verify_and_send_email(self.rfq_number.get(), other_attachment=list(self.other_attachments.get(0, tk.END)), item_id=self.item_id.get(), qty_req=self.item_qty.get()))
        send_mail_button.grid(row=5, column=0)

    def verify_and_send_email(self, rfq_number, other_attachment=[], item_id=None, qty_req=None):
        if rfq_number:
            send_mail(rfq_number=rfq_number, other_attachment=other_attachment)
            messagebox.showinfo("Success", "Email sent successfully")
            self.rfq_number.delete(0, tk.END)
            self.other_attachments.delete(0, tk.END)
        elif item_id and qty_req:
            send_mail(item_id=item_id, qty_req=qty_req, other_attachment=other_attachment)
            messagebox.showinfo("Success", "Email sent successfully")
            self.rfq_number.delete(0, tk.END)
            self.other_attachments.delete(0, tk.END)
        else:
            messagebox.showerror("Error", "Please enter RFQ number/ ItemID")
            self.other_attachments.delete(0, tk.END)
    
    def browse_files_parts_requested(self, filetype: str, list_box):
        """ Browse button for Part requested section, filetype only accepts -> "All files", "Excel files" """
        if filetype == "Excel files":
            param = (filetype, "*.xlsx;*.xls")
        else:
            param = (filetype, "*.*")

        try:
            self.filepaths = [filepath for filepath in filedialog.askopenfilenames(title="Select Files", filetypes=(param,))]

            # entering all file paths in the listbox
            list_box.delete(0, tk.END)
            for path in self.filepaths:
                list_box.insert(0, path)

        except FileNotFoundError as e:
            print(f"Error during file browse: {e}")
            messagebox.showerror("File Browse Error", "An error occurred during file selection. Please try again.")

if __name__ == "__main__":
    app = EmailGui()
    app.mainloop()
