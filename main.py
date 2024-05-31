from src.general_class import TableManger
import openpyxl
import os

def get_item_pks(rfq_pk):
    rfq_line_table = TableManger("RequestForQuoteLine")
    quote_assembly_table = TableManger("QuoteAssembly")
    quote_fk = rfq_line_table.get("QuoteFK", RequestForQuoteFK=rfq_pk)
    all_item_pks = []
    item_fks = []
    if quote_fk:
        for fk in quote_fk:
            item_pks = quote_assembly_table.get("ItemFK", QuoteFK=fk[0])            
            if item_pks:
                all_item_pks.extend(item_pks)
        for pk in all_item_pks:
            if pk[0] is not None:
                item_fks.append(pk[0])
    return item_fks

def get_item_dict(item_fks):
    item_table = TableManger("Item")
    item_dict = {}
    
    for fk in item_fks:
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
            email_category = "mat-al"  # Default email_category to the category if not specified
        
        item_dict[itempk]["Category"] = category
        item_dict[itempk]["EmailCategory"] = email_category
 
    return item_dict


def create_excel(filepath, item_dict, rfq_number):
    # Load the existing Excel file
    workbook = openpyxl.load_workbook(filepath)
    new_dir = os.path.join("RFQ_Excel", str(rfq_number))
    new_filepath = os.path.join(new_dir, "_".join(filepath.split("_")[2:]))
    
    # Ensure the directory exists
    os.makedirs(new_dir, exist_ok=True)
    sheet = workbook.active
    sheet.delete_rows(3, sheet.max_row - 2)
    row_idx = 3
    for key, values in item_dict.items():
        for key1, val in values.items():
            if filepath=="RFQ_template_hardware.xlsx":
                if key1=="EmailCategory" and val=="hardware":
                    
                    sheet.cell(row=row_idx, column=1).value = values['PartNumber']
                    sheet.cell(row=row_idx, column=2).value = values['Description']
                    sheet.cell(row=row_idx, column=6).value = values['PartLength']
                    sheet.cell(row=row_idx, column=5).value = values['PartWidth']
                    sheet.cell(row=row_idx, column=4).value = values['Thickness']
                    sheet.cell(row=row_idx, column=7).value = values['Comment']
                    row_idx += 1
            elif filepath=="RFQ_template_fin.xlsx":
                if key1=="EmailCategory" and val=="fin":
                    sheet.cell(row=row_idx, column=1).value = values['PartNumber']
                    sheet.cell(row=row_idx, column=2).value = values['Description']
                    sheet.cell(row=row_idx, column=6).value = values['PartLength']
                    sheet.cell(row=row_idx, column=5).value = values['PartWidth']
                    sheet.cell(row=row_idx, column=4).value = values['Thickness']
                    sheet.cell(row=row_idx, column=7).value = values['Comment']
                    row_idx += 1
            elif filepath=="RFQ_template_mat_al.xlsx":
                if key1=="EmailCategory" and val=="mat-al":
                    sheet.cell(row=row_idx, column=1).value = values['PartNumber']
                    sheet.cell(row=row_idx, column=2).value = values['Description']
                    sheet.cell(row=row_idx, column=6).value = values['StockLength']
                    sheet.cell(row=row_idx, column=5).value = values['StockWidth']
                    sheet.cell(row=row_idx, column=4).value = values['Thickness']
                    sheet.cell(row=row_idx, column=7).value = values['Comment']
                    row_idx += 1
            elif filepath=="RFQ_template_mat_steel.xlsx":
                if key1=="EmailCategory" and val=="mat-steel":
                    sheet.cell(row=row_idx, column=1).value = values['PartNumber']
                    sheet.cell(row=row_idx, column=2).value = values['Description']
                    sheet.cell(row=row_idx, column=6).value = values['StockLength']
                    sheet.cell(row=row_idx, column=5).value = values['StockWidth']
                    sheet.cell(row=row_idx, column=4).value = values['Thickness']
                    sheet.cell(row=row_idx, column=7).value = values['Comment']
                    row_idx += 1
            elif filepath=="RFQ_template_ht.xlsx":
                if key1=="EmailCategory" and val=="ht":
                    sheet.cell(row=row_idx, column=1).value = values['PartNumber']
                    sheet.cell(row=row_idx, column=2).value = values['Description']
                    sheet.cell(row=row_idx, column=6).value = values['PartLength']
                    sheet.cell(row=row_idx, column=5).value = values['PartWidth']
                    sheet.cell(row=row_idx, column=4).value = values['Thickness']
                    sheet.cell(row=row_idx, column=7).value = values['Comment']
                    row_idx += 1

    workbook.save(new_filepath)


# TODO: log creation of excel sheets and folders


def send_all_emails(filepath_folder_of_excel_sheets: str) -> None:
    """Sends all emails to the IDs listed in the email sheet for each category"""
    pass

def create_excel_sheets(rfq_number):
    item_dict = get_item_dict(get_item_pks(rfq_number))
    main_dict = sort_items_in_groups(item_dict)
    category_list = []
    for key, value in main_dict.items():
        for key1, value1 in value.items():
            if key1=="EmailCategory" and value1 not in category_list:
                category_list.append(value1)
    for category in category_list:
        if category == "hardware":
            filepath = "RFQ_template_hardware.xlsx"
            create_excel(filepath, main_dict, rfq_number)
        elif category == "ht":
            filepath = "RFQ_template_ht.xlsx"
            create_excel(filepath, main_dict, rfq_number)
        elif category == "mat-steel":
            filepath = "RFQ_template_mat_steel.xlsx"
            create_excel(filepath, main_dict, rfq_number)
        elif category == "mat-al":
            filepath = "RFQ_template_mat_al.xlsx"
            create_excel(filepath, main_dict, rfq_number)
        elif category== "fin":
            filepath = "RFQ_template_fin.xlsx"
            create_excel(filepath, main_dict, rfq_number)
        




if __name__=="__main__":
    create_excel_sheets(5995)

 