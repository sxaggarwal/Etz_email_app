import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from src.helper import send_mail, get_items_dict, get_rfq_pk


class EmailGui(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Send Email App")
        self.geometry("310x400")
        self.make_combobox()
        self.item_dict = get_items_dict()
        self.rfq_pk = get_rfq_pk()

    def make_combobox(self):
        """Main GUI of the email app"""
        tk.Label(self, text="Select Type: ").grid(row=0, column=0)
        self.type_select_box = ttk.Combobox(
            self, values=["Item", "RFQ"], state="normal"
        )
        self.type_select_box.grid(row=1, column=0)

        tk.Label(self, text="Search for RFQ/ Item").grid(row=2, column=0)
        self.rfq_or_item_search = tk.Entry(self, width=20)
        self.rfq_or_item_search.grid(row=3, column=0)
        self.rfq_or_item_search.bind("<Return>", self.search_documents)

        search_rfq_or_item_button = tk.Button(
            self, text="Search", command=self.search_documents
        )
        search_rfq_or_item_button.grid(row=4, column=0)

        tk.Label(self, text="Select RFQ/ Item").grid(row=5, column=0)
        self.search_result_box = tk.Listbox(
            self, height=4, width=30, exportselection=False
        )
        self.search_result_box.grid(row=6, column=0)

        tk.Label(self, text="Enter Item Qty Req ").grid(row=10, column=0)
        self.item_qty = tk.Entry(self, width=20)
        self.item_qty.grid(row=11, column=0)

        tk.Label(self, text="Other Attachments:").grid(row=7, column=0)
        self.other_attachments = tk.Listbox(self, height=2, width=50)
        self.other_attachments.grid(row=8, column=0)

        browse_button_part_list = tk.Button(
            self,
            text="Browse Files",
            command=lambda: self.browse_files_parts_requested(
                "All files", self.other_attachments
            ),
        )
        browse_button_part_list.grid(row=9, column=0)

        send_mail_button = tk.Button(
            self,
            text="Send Email",
            command=lambda: self.verify_and_send_email(
                self.get_pk("RFQ"),
                other_attachment=list(self.other_attachments.get(0, tk.END)),
                item_id=self.get_pk("Item"),
                qty_req=self.item_qty.get(),
            ),
        )
        send_mail_button.grid(row=12, column=0)

    def verify_and_send_email(
        self, rfq_number, other_attachment=[], item_id=None, qty_req=None
    ):
        """Verifies if everything if properly filled in the GUI and then calls the send_mail function to send the email"""
        selected_type = self.type_select_box.get()

        if selected_type == "RFQ" and rfq_number:
            send_mail(rfq_number=rfq_number, other_attachment=other_attachment)
            messagebox.showinfo("Success", "Email sent successfully")
            # self.rfq_number.delete(0, tk.END)
            self.other_attachments.delete(0, tk.END)
            self.search_result_box.delete(0, tk.END)
            self.item_qty.delete(0, tk.END)
            self.rfq_or_item_search.delete(0, tk.END)
        elif selected_type == "Item" and item_id and qty_req:
            send_mail(
                item_id=item_id, qty_req=qty_req, other_attachment=other_attachment
            )
            messagebox.showinfo("Success", "Email sent successfully")
            # self.rfq_number.delete(0, tk.END)
            self.other_attachments.delete(0, tk.END)
            self.search_result_box.delete(0, tk.END)
            self.item_qty.delete(0, tk.END)
            self.rfq_or_item_search.delete(0, tk.END)
        else:
            messagebox.showerror("Error", "Please enter RFQ number/ ItemID")
            self.other_attachments.delete(0, tk.END)
            self.search_result_box.delete(0, tk.END)
            self.item_qty.delete(0, tk.END)
            self.rfq_or_item_search.delete(0, tk.END)

    def get_pk(self, type1):
        """Gets the primary key of the selected item in the combobox"""
        selected_index = self.search_result_box.curselection()[0]
        selected_item = self.search_result_box.get(selected_index)
        selected_type = self.type_select_box.get()
        pk = None

        if type1 == "Item" and selected_type == "Item":
            pk = selected_item.split("-")
            if pk[0]:
                pk = pk[0].strip()
            # else:
            #     pk = None
        elif type1 == "RFQ" and selected_type == "RFQ":
            pk = selected_item.split(":")
            if pk[1]:
                pk = pk[1].strip()
            # else:
            #     pk = None
        return pk

    def browse_files_parts_requested(self, filetype: str, list_box):
        """Browse button for Part requested section, filetype only accepts -> "All files", "Excel files" """
        if filetype == "Excel files":
            param = (filetype, "*.xlsx;*.xls")
        else:
            param = (filetype, "*.*")

        try:
            self.filepaths = [
                filepath
                for filepath in filedialog.askopenfilenames(
                    title="Select Files", filetypes=(param,)
                )
            ]

            # entering all file paths in the listbox
            list_box.delete(0, tk.END)
            for path in self.filepaths:
                list_box.insert(0, path)

        except FileNotFoundError as e:
            print(f"Error during file browse: {e}")
            messagebox.showerror(
                "File Browse Error",
                "An error occurred during file selection. Please try again.",
            )

    def search_documents(self, event=None):
        selected_type = self.type_select_box.get()
        user_search = self.rfq_or_item_search.get()
        self.search_result_box.delete(0, tk.END)
        if selected_type == "Item":
            for key, value in self.item_dict.items():
                if value is not None and value.startswith(user_search):
                    self.search_result_box.insert(tk.END, f"{key} - {value}")
        elif selected_type == "RFQ":
            for pk in self.rfq_pk:
                if str(pk[0]).startswith(user_search):
                    self.search_result_box.insert(tk.END, f"RFQ Number: {str(pk[0])}")


if __name__ == "__main__":
    app = EmailGui()
    app.mainloop()
