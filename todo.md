[ ] Integration with Email App.
    [ ] Read rest of the send email function
    [ ] change the email list retreival process


# Flow
INPUT - RFQ
Get all items from rfq
get all item commodity codes. <item> : "commodity_code": <code>
    add the commodity code to set.
get all item values from the item table (use get_item_dict)
Build excel sheet for each code.

# GUI Flow: 
-> User presses the send mail button
    - get all items and their commodities
    - show them on a new GUI
    NOTE: notify of errors or commodities not filled out in items.
