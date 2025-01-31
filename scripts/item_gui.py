import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, List

class PopupWindow(tk.Toplevel):
    def __init__(self, parent, data_dict: Dict[str, List[str]]):
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
        self.email_listbox = tk.Listbox(container)
        self.email_listbox.pack(side="right", fill="both", expand=True, padx=5)

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
            self.email_listbox.delete(0, tk.END)
            emails = self.data_dict.get(self.selected_code, ["No Emails Found"])
            for email in emails:
                self.email_listbox.insert(tk.END, email)

    def add_email(self):
        """Add a new email to the selected code."""
        if not self.selected_code:
            messagebox.showwarning("Error", "No code selected!")
            return

        new_email = self.email_entry.get().strip()
        if new_email:
            # Add email to dictionary
            if self.selected_code in self.data_dict:
                if new_email not in self.data_dict[self.selected_code]:
                    self.data_dict[self.selected_code].append(new_email)
                    self.email_listbox.insert(tk.END, new_email)
                    self.email_entry.delete(0, tk.END)
                else:
                    messagebox.showwarning("Error", "Email already exists!")
            else:
                messagebox.showwarning("Error", "Invalid Code Selection!")
        else:
            messagebox.showwarning("Error", "Enter a valid email!")

    def update_email(self):
        """Update the selected email."""
        if not self.selected_code:
            messagebox.showwarning("Error", "No code selected!")
            return

        selected_email_index = self.email_listbox.curselection()
        if not selected_email_index:
            messagebox.showwarning("Error", "No email selected!")
            return

        new_email = self.email_entry.get().strip()
        if new_email:
            # Get selected email and replace it
            selected_email = self.email_listbox.get(selected_email_index)
            if selected_email in self.data_dict[self.selected_code]:
                email_index = self.data_dict[self.selected_code].index(selected_email)
                self.data_dict[self.selected_code][email_index] = new_email
                self.email_listbox.delete(selected_email_index)
                self.email_listbox.insert(selected_email_index, new_email)
                self.email_entry.delete(0, tk.END)
            else:
                messagebox.showwarning("Error", "Email not found in database!")
        else:
            messagebox.showwarning("Error", "Enter a valid email!")

    def delete_email(self):
        """Delete the selected email."""
        if not self.selected_code:
            messagebox.showwarning("Error", "No code selected!")
            return

        selected_email_index = self.email_listbox.curselection()
        if not selected_email_index:
            messagebox.showwarning("Error", "No email selected!")
            return

        selected_email = self.email_listbox.get(selected_email_index)

        # Confirm deletion
        confirm = messagebox.askyesno("Confirm", f"Delete email '{selected_email}'?")
        if confirm:
            # Remove email from dictionary
            if selected_email in self.data_dict[self.selected_code]:
                self.data_dict[self.selected_code].remove(selected_email)
                self.email_listbox.delete(selected_email_index)
                self.email_entry.delete(0, tk.END)

