from src.general_class import TableManger
import openpyxl
import os
import win32com.client as win32
import pythoncom
import time
from  tkinter import messagebox, ttk, simpledialog
import subprocess
import tkinter as tk
from tkinter.scrolledtext import ScrolledText

def get_item_pks(rfq_pk):
    rfq_line_table = TableManger("RequestForQuoteLine")
    quote_assembly_table = TableManger("QuoteAssembly")
    quote_fk = rfq_line_table.get("QuoteFK", RequestForQuoteFK=rfq_pk)
    all_item_pks = []
    item_fks_dict = {}
    if quote_fk:
        for fk in quote_fk:
            item_pks = quote_assembly_table.get("ItemFK", "QuantityRequired", QuoteFK=fk[0])         
            if item_pks:
                all_item_pks.extend(item_pks)
        for pk in all_item_pks:
            if pk[0] is not None:
               item_fks_dict[pk[0]] = pk[1]
    return item_fks_dict

def create_single_item_dict(item_id, qty_req):
    item_dict = {}
    item_dict[item_id] = qty_req
    return item_dict

def get_item_dict(item_fks_dict):
    item_table = TableManger("Item")
    item_dict = {}
    
    for fk, value in item_fks_dict.items():
        details = item_table.get(
            "PartNumber", "Description", "ItemTypeFK", "PartLength",
            "PartWidth", "Thickness", "StockLength", "StockWidth", "Comment",
            ItemPK=fk
        )
        if details:
            columns = ["PartNumber", "Description", "ItemTypeFK", "PartLength",
                       "PartWidth", "Thickness", "StockLength", "StockWidth", "Comment"]
            item_details = details[0]  
            item_dict[fk] = {column: value for column, value in zip(columns, item_details)}
            item_dict[fk]["QuantityRequired"] = value
    
    return item_dict

def get_email_groups():
    """ Returns all emails that belong to the following groups
        'material', 'hardware', 'op', 'tooling'. 
        {'material': <values>, 'op': <values> ...}
     """
    party_buyer_table = TableManger("PartyBuyer") 
    party_table = TableManger("Party")
    parties = {                              # name : partyFK (we get buyerfk from this)
        "fin" : 3755,
        "mat-al": 3744,
        "mat-steel": 3747,
        "ht": 3764,
        'hardware': 3867,
    }
    email_dict = {}
    for key, partyfk in parties.items():
        buyerfk = party_buyer_table.get("BuyerFK", PartyFK=partyfk)
        email_dict[key] = [] 
        for fk in buyerfk:
            email = party_table.get("Email", PartyPK=fk[0])[0][0]
            email_dict[key].append(email)
 
    return email_dict
 
 
# the categories in the below function are meant to be directly used in the email_dict returned in get_email_groups
# the category that is appended to the item dict will match the key in email dict returned in get_email_groups
 

def sort_items_in_groups(item_dict: dict):
    """Identifies item type and then appends the category as a key, value pair to the item_dict."""
    categories = ["standard", "mat", "hardware", "msc", "fin", "kit", "tooling"]
    
    for itempk, column_values in item_dict.items():
        item_type_index = item_dict[itempk]["ItemTypeFK"] - 1  # to get the corresponding index in categories
        category = categories[item_type_index]
        
        part_number_split = item_dict[itempk]["PartNumber"].lower().split()
        email_category = None  # Initialize email_category to avoid reference before assignment

        if category == "fin":
            if "ht" in part_number_split:
                email_category = "ht"
            else:
                email_category = "fin"
        elif category in ["hardware", "tooling"]:
            email_category = "hardware"
        elif category == "mat":
            if "al" in part_number_split:
                email_category = "mat-al"
            elif "steel" in part_number_split or "st" in part_number_split:
                email_category = "mat-steel"
            else:
                email_category = "mat-al"
        else:
            email_category = "mat-al"  # Default email_category to the category if not specified
        
        item_dict[itempk]["Category"] = category
        item_dict[itempk]["EmailCategory"] = email_category
 
    return item_dict


def create_excel(filepath, item_dict, rfq_number):
    workbook = openpyxl.load_workbook(filepath)
    # new_dir = os.path.join(r".\RFQ_Excel", str(rfq_number))
    new_dir = os.path.abspath(os.path.join(".", "RFQ_Excel", str(rfq_number)))
    new_filepath = os.path.join(new_dir, "_".join(filepath.split("_")[2:]))
    os.makedirs(new_dir, exist_ok=True)
    sheet = workbook.active
    sheet.delete_rows(3, sheet.max_row - 2)
    row_idx = 3
    for key, values in item_dict.items():
        if filepath == f"RFQ_template_{values['EmailCategory']}.xlsx":
            sheet.cell(row=row_idx, column=1).value = values['PartNumber']
            sheet.cell(row=row_idx, column=2).value = values['Description']
            sheet.cell(row=row_idx, column=6).value = values['StockLength'] if values['EmailCategory'] in ['mat-al', 'mat-steel'] else values['PartLength']
            sheet.cell(row=row_idx, column=5).value = values['StockWidth'] if values['EmailCategory'] == 'mat-al' or values['Category']=='mat-steel' else values['PartWidth']
            sheet.cell(row=row_idx, column=4).value = values['Thickness']
            sheet.cell(row=row_idx, column=7).value = values['Comment']
            sheet.cell(row=row_idx, column=3).value = values['QuantityRequired']
            #Insert here
            row_idx += 1
    workbook.save(new_filepath)
    return new_filepath


def send_all_emails(filepath_folder_of_excel_sheets: str) -> None:
    """Sends all emails to the IDs listed in the email sheet for each category"""
    pass


def create_excel_sheets(rfq_number=None, item_id=None, qty_req=None):
    if rfq_number:
        item_dict = get_item_dict(get_item_pks(rfq_number))
    elif item_id:
        item_dict = get_item_dict(create_single_item_dict(item_id, qty_req)) #put a dict
    main_dict = sort_items_in_groups(item_dict)
    category_list = []
    excel_path_list = []
    for key, value in main_dict.items():
        for key1, value1 in value.items():
            if key1 == "EmailCategory" and value1 not in category_list:
                category_list.append(value1)
    for category in category_list:
        filepath = f"RFQ_template_{category}.xlsx"
        excel_path = create_excel(filepath, main_dict, rfq_number)
        excel_path_list.append(excel_path)
    return excel_path_list


def send_outlook_email(excel_path, email_list, subject, email_body, other_attachment=[], cc_email=None):
    pythoncom.CoInitialize()
    try:
        outlook = win32.Dispatch("outlook.application")
        print("Outlook email creation started")

        for email in email_list:
            try:
                mail = outlook.CreateItem(0)  # Create a new mail item for each recipient
                mail.Subject = subject
                mail.Body = email_body
                mail.To = email
                mail.CC = cc_email

                abs_excel_path = os.path.abspath(excel_path)
                if os.path.isfile(abs_excel_path):
                    mail.Attachments.Add(abs_excel_path)
                else:
                    print(f"Attachment not found: {abs_excel_path}")

                for attachment in other_attachment:
                    abs_attachment_path = os.path.abspath(attachment)
                    if os.path.isfile(abs_attachment_path):
                        mail.Attachments.Add(abs_attachment_path)
                    else:
                        print(f"Additional attachment not found: {abs_attachment_path}")

                mail.Send()
                print(f"Email sent to {email}")
                time.sleep(1)  # Adding a delay to avoid rate limits or other issues
            except Exception as e:
                print(f"An error occurred while sending email to {email}: {e}")
                # Retry logic can be added here if needed
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        pythoncom.CoUninitialize()

class EmailBodyDialog(simpledialog.Dialog):
    def __init__(self, parent, initial_body):
        self.initial_body = initial_body
        super().__init__(parent, title = "Edit Email Body")

    def body(self, master):
        self.title("Edit Email Body")
        self.geometry("600x600")
        self.resizable(True, True)
        tk.Label(master, text="Edit the email body below:").pack(pady=5)
        self.text = ScrolledText(master, wrap=tk.WORD)
        self.text.pack(expand=True, fill=tk.BOTH)
        self.text.insert(tk.END, self.initial_body)
        return self.text

    def apply(self):
        self.result = self.text.get("1.0", tk.END)

def email_body_template(root): #NOTE: In future we can somehow connect this to the database and fill info automatically of the user.
    initial_body = ("Dear Supplier, \n"
            "Hope you are doing well,\n"
            "This email serves as a formal request for a quotation on pricing and lead time for the items listed in the attached Excel spreadsheet.\n"
            "For your convenience, we have included several columns in the spreadsheet. While completion of all columns is not mandatory, we kindly request that you fill out those you deem most relevant to accurately reflecting your pricing and lead time for each item.\n"
            "Please note that to ensure a proper evaluation of your offer, a completed spreadsheet is necessary.\n"
            "We appreciate your prompt attention to this matter. Should you require any clarification regarding the listed items, please do not hesitate to contact us.\n"
            "\n"
            "Thankyou.\n"
            "<your_name>\n"
            "<your_designation>\n"
            "Etezazi Industries Inc.\n"
            "2101 E. 21st St North\n"
            "Wichita, KS 67214\n"
            "(316)-831 9937 Office EXT <your_ext>\n"
            "Email: <your_email>\n"
            "Website: www.etezazi-industries.com\n"
            "AS9100D and ISO9001:2015 and ITAR Registered Company\n")
    dialog = EmailBodyDialog(root, initial_body)
    # root.wait_window(dialog)
    return dialog.result.strip() if dialog.result else initial_body 

def send_mail(rfq_number=None, other_attachment = [], item_id=None, qty_req=None):
    
    email_dict = get_email_groups()    #NOTE: Uncomment this when needed
    # NOTE: below is the test email id's, and you can comment that and uncomment above for real supplier IDS
    # email_dict = {
    #     'mat-al': ['shubham.aggarwal@etezazicorps.com'],
    #     'fin': ['shubham.smvit@gmail.com'],
    #     'hardware': ['shubham.smvit@gmail.com'],
    # }
    if rfq_number:
        subject = f"RFQ - {rfq_number}"
        excel_path_list = create_excel_sheets(rfq_number=rfq_number)
    else:
        subject = f"RFQ - {item_id}"
        excel_path_list = create_excel_sheets(item_id=item_id, qty_req=qty_req)
    root = tk.Tk()
    root.withdraw()
    email_body = email_body_template(root)
    for excel_path in excel_path_list:
        response = messagebox.askyesno("View/Edit Excel", f"Do you want to View and Edit the Excel file: {excel_path}")
        if response:
            subprocess.Popen(['start', excel_path], shell=True)
            messagebox.showinfo("Info", f"Please edit and save the file: {excel_path}")
            while True:
                edit_complete = messagebox.askyesno("Edit Complete", "Have you finished editing and saving the file?")
                if edit_complete:
                    break
                else:
                    messagebox.showinfo("Info", "Please complete your edits and save the file before sending the email.")
        else:
            pass
        excel_filename = os.path.basename(excel_path)
        for key, values in email_dict.items():
            if key in excel_filename:
                # email_list_str = "\n".join(values)
                # new_email_list_str = simpledialog.askstring("Edit Email IDs", f"Current email IDs for {key}:\n{email_list_str}\n\nEdit email IDs (separated by commas):", initialvalue=", ".join(values))
                # if new_email_list_str:
                #     new_email_list = [email.strip() for email in new_email_list_str.split(",")]
                # else:
                #     new_email_list = values
                new_email_list = get_email_input(root, f"Edit Email IDs for {key}", values)
                if new_email_list is not None:
                    send_outlook_email(excel_path, new_email_list, subject, email_body, other_attachment=other_attachment, cc_email="quote@etezazicorps.com")
                # send_outlook_email(excel_path, new_email_list, subject, other_attachment=other_attachment)
    root.destroy()

def get_email_input(parent, title, initial_emails):
    dialog = EmailDialog(parent, title, initial_emails)
    parent.wait_window(dialog.top)
    return dialog.result

class EmailDialog:
    def __init__(self, parent, title, initial_emails):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("400x300")

        self.label = tk.Label(self.top, text="Edit email IDs:")
        self.label.pack(pady=10)

        self.listbox = tk.Listbox(self.top, selectmode=tk.SINGLE)
        self.listbox.pack(expand=True, fill='both', padx=10, pady=10)
        for email in initial_emails:
            self.listbox.insert(tk.END, email)

        self.entry = tk.Entry(self.top)
        self.entry.pack(pady=5)

        self.add_button = ttk.Button(self.top, text="Add Email", command=self.add_email)
        self.add_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.remove_button = ttk.Button(self.top, text="Remove Selected Email", command=self.remove_email)
        self.remove_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.ok_button = ttk.Button(self.top, text="OK", command=self.on_ok)
        self.ok_button.pack(side=tk.RIGHT, padx=5, pady=5)

        self.result = None

    def add_email(self):
        email = self.entry.get().strip()
        if email and email not in self.listbox.get(0, tk.END):
            self.listbox.insert(tk.END, email)
            self.entry.delete(0, tk.END)

    def remove_email(self):
        selected_idx = self.listbox.curselection()
        if selected_idx:
            self.listbox.delete(selected_idx)

    def on_ok(self):
        self.result = list(self.listbox.get(0, tk.END))
        self.top.destroy()
        