import tkinter as tk
from mt_api.general_class import TableManger
from mt_api.connection import get_connection
from tkinter import ttk, messagebox
from typing import Dict, List, Tuple


class PopupWindow(tk.Toplevel):
    def __init__(self, parent, data_dict: Dict[str, List[Tuple[str, str]]]):
        super().__init__(parent)
        self.title("Manage Emails")
        self.geometry("500x400")

        # Store data dictionary
        self.data_dict = data_dict
        self.selected_code = None  # Track currently selected code

        # Create a frame for layout
        container = tk.Frame(self)
        container.pack(fill="both", expand=True, padx=10, pady=10)

        # Left Side: Treeview for Codes
        self.tree = ttk.Treeview(container, columns=("Code",), show="headings", selectmode="browse")
        self.tree.heading("Code", text="Code")
        self.tree.column("Code", width=150, anchor="center")

        # Insert codes into the Treeview
        for code in self.data_dict.keys():
            self.tree.insert("", "end", values=(code,))

        self.tree.pack(side="left", fill="both", expand=True, padx=5)

        # Bind selection event
        self.tree.bind("<<TreeviewSelect>>", self.on_code_selected)

        # Right Side: Listbox for Emails
        self.email_tree = ttk.Treeview(container, columns=("Name", "Email"), show="headings", selectmode="browse")
        self.email_tree.heading("Name", text="Name")
        self.email_tree.heading("Email", text="Email")
        self.email_tree.column("Name", width=150, anchor="w")
        self.email_tree.column("Email", width=250, anchor="w")
        self.email_tree.pack(side="right", fill="both", expand=True, padx=5)

        # CRUD Frame for Emails
        email_frame = tk.Frame(self)
        email_frame.pack(fill="x", padx=10, pady=5)

        # Input Field for Email
        ttk.Label(email_frame, text="Email:").grid(row=0, column=0, sticky="w")
        self.email_entry = ttk.Entry(email_frame)
        self.email_entry.grid(row=0, column=1, padx=5, pady=2, sticky="ew")

        # CRUD Buttons
        ttk.Button(email_frame, text="Add Email", command=self.add_email).grid(row=1, column=0, pady=5, sticky="ew")
        ttk.Button(email_frame, text="Update Email", command=self.update_email).grid(row=1, column=1, pady=5, sticky="ew")
        ttk.Button(email_frame, text="Delete Email", command=self.delete_email).grid(row=1, column=2, pady=5, sticky="ew")

    def on_code_selected(self, event):
        """Display emails when a code is selected."""
        selected_item = self.tree.selection()
        if not selected_item:
            return  # No selection, do nothing

        item_values = self.tree.item(selected_item[0], "values")
        if item_values:
            self.selected_code = item_values[0]

            # Clear previous email list and show emails
            for row in self.email_tree.get_children():
                self.email_tree.delete(row)

            # Get emails for selected code
            emails = self.data_dict.get(self.selected_code, [])

            # Insert into Treeview
            # DEBUG: some error in indexing
            for name, email in emails:
                self.email_tree.insert("", "end", values=(name, email))

            # If no emails are found, show a placeholder row
            if not emails:
                self.email_tree.insert("", "end", values=("No Name", "No Email Found"))


    def add_email(self):
        """Add a new email to the selected code."""
        if not self.selected_code:
            messagebox.showwarning("Error", "No code selected!")
            return

        new_email = self.email_entry.get().strip()
        if new_email:
            customer_group_table = TableManger("CustomerGroup")
            customer_group_pk = customer_group_table.get("CustomerGroupPK", Code=self.selected_code)
            if not customer_group_pk:
                raise ValueError(f"Table returned no values for code: {self.selected_code}")

            # selection popup to find party PKs
            PartySelection(self, self._add_email_callback)




        # TODO: 
        # Find the code's customergroup pk from code
        # Find partyfk - user input.
        # party customer group add - partyfk, customergroupfk.
        # update our email list values as well.

    @staticmethod
    def _add_email_callback():
        party_customer_group_table = TableManger("PartyCustomerGroup")



    def update_email(self):
        """Update the selected email."""
        # if not self.selected_code:
        #     messagebox.showwarning("Error", "No code selected!")
        #     return
        #
        # selected_email_index = self.email_listbox.curselection()
        # if not selected_email_index:
        #     messagebox.showwarning("Error", "No email selected!")
        #     return
        #
        # new_email = self.email_entry.get().strip()
        # if new_email:
        #     # Get selected email and replace it
        #     selected_email = self.email_listbox.get(selected_email_index)
        #     if selected_email in self.data_dict[self.selected_code]:
        #         email_index = self.data_dict[self.selected_code].index(selected_email)
        #         self.data_dict[self.selected_code][email_index] = new_email
        #         self.email_listbox.delete(selected_email_index)
        #         self.email_listbox.insert(selected_email_index, new_email)
        #         self.email_entry.delete(0, tk.END)
        #     else:
        #         messagebox.showwarning("Error", "Email not found in database!")
        # else:
        #     messagebox.showwarning("Error", "Enter a valid email!")
        pass

    def delete_email(self):
        """Delete the selected email."""
        # if not self.selected_code:
        #     messagebox.showwarning("Error", "No code selected!")
        #     return
        #
        # selected_email_index = self.email_listbox.curselection()
        # if not selected_email_index:
        #     messagebox.showwarning("Error", "No email selected!")
        #     return
        #
        # selected_email = self.email_listbox.get(selected_email_index)
        #
        # # Confirm deletion
        # confirm = messagebox.askyesno("Confirm", f"Delete email '{selected_email}'?")
        # if confirm:
        #     # Remove email from dictionary
        #     if selected_email in self.data_dict[self.selected_code]:
        #         self.data_dict[self.selected_code].remove(selected_email)
        #         self.email_listbox.delete(selected_email_index)
        #         self.email_entry.delete(0, tk.END)
        pass



class PartySelection(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Party Selection")
        self.geometry("800x400")  # Adjusted for wider display
        self.create_widgets()

    def create_widgets(self):
        """Create the search bar, button, and results table."""
        # Frame for search bar and button
        search_frame = tk.Frame(self)
        search_frame.pack(fill="x", padx=10, pady=5)

        # Search bar
        ttk.Label(search_frame, text="Search:").pack(side="left", padx=5)
        self.search_entry = ttk.Entry(search_frame)
        self.search_entry.pack(side="left", fill="x", expand=True, padx=5)

        # Search button
        self.search_button = ttk.Button(search_frame, text="Search", command=self.perform_search)
        self.search_button.pack(side="left", padx=5)

        # Frame for Treeview and Scrollbars
        results_frame = tk.Frame(self)
        results_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Treeview (8 columns)
        self.results_tree = ttk.Treeview(
            results_frame, 
            columns=("Col1", "Col2", "Col3", "Col4", "Col5", "Col6", "Col7", "Col8"), 
            show="headings"
        )

        # Define column headers
        for i in range(1, 9):
            col_id = f"Col{i}"
            self.results_tree.heading(col_id, text=f"Column {i}")
            self.results_tree.column(col_id, anchor="w", width=120)  # Adjust width as needed

        # Scrollbars
        vsb = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_tree.yview)
        hsb = ttk.Scrollbar(results_frame, orient="horizontal", command=self.results_tree.xview)

        self.results_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Layout
        self.results_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")  # Vertical scrollbar
        hsb.pack(side="bottom", fill="x")  # Horizontal scrollbar

    def perform_search(self):
        """This function should handle search logic."""
        results = self.search()

        # Clear previous results
        for row in self.results_tree.get_children():
            self.results_tree.delete(row)

        # Insert search results into Treeview (assuming 8 values per row)
        for row in results:
            if len(row) == 8:  # Ensure row contains exactly 8 elements
                self.results_tree.insert("", "end", values=row)

    def search(self) -> List[List[str]]:
        """Placeholder search function. Implement your logic here."""
        return [
            ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1"],
            ["A2", "B2", "C2", "D2", "E2", "F2", "G2", "H2"],
            ["A3", "B3", "C3", "D3", "E3", "F3", "G3", "H3"],
        ]  # Example return data (8 columns per row)


def ser():
    search_value = "Metal"
    if not search_value:
        messagebox.showerror(message="Search value cannot be empty")

    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(f"SELECT Name, PartyPK FROM Party WHERE Name LIKE '{search_value}%'")
        results = cursor.fetchall()

        print(results)

