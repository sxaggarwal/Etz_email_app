from typing import Any, List, Tuple, Dict
from pprint import pprint
from xlwings.main import Table
from mt_api.general_class import TableManger
from mt_api.base_logger import getlogger
from scripts.helper import get_item_dict


LOGGER = getlogger("mt_commodity_script")


# NOTE: we can use the get_item_pks in helper to get all the itempks first before we run this script for each item.


class Controller:
    def __init__(self) -> None:
        pass

    def get_commodity_from_item(self, itempk: int) -> str:
        """
        Returns the Commodity Code associated with the Item.

        :param itempk int: Primay key of Item from Item Table.
        :return: Commodity Code by referencing the Commodity Table
        :raises ValueError: If CommodityFK is empty in Item Table
        """
        item_table = TableManger("Item")
        commodity_fk = item_table.get("CommodityFK", ItemPK=itempk)

        if not commodity_fk[0][0]:
            raise ValueError(f"Item with ItemPK: {itempk} does not have a commodity in Mie Trak")

        commodity_table = TableManger("Commodity")
        code = commodity_table.get("Code", CommodityPK=commodity_fk[0][0])

        return code[0][0]

    def search_for_rfq(self, rfq_number):
        request_for_quote_table = TableManger("RequestForQuote")
        result = request_for_quote_table.get("RequestForQuotePK", RequestForQuotePK=rfq_number)
        return result[0][0]

    def get_all_line_items_for_rfq(self, rfq_pk) -> Dict[int, Dict[str, Any]]:
        """
        Gets the ItemPKs of all items in the RFQ and aggregates the quantities needed.

        :param rfq_pk: PK in the RFQ table in MT.
        :return: {<itemfk>: <total qty>}
        :raises ValueError: When no data is returned or no qty is returned from one of them items (line items).
        """

        rfq_line_table = TableManger("RequestForQuoteLine")
        quote_assembly_table = TableManger("QuoteAssembly")

        # Retrieve the list of (QuoteFK, Quantity) for the given RFQ
        quote_fk_list = rfq_line_table.get("QuoteFK", "Quantity", RequestForQuoteFK=rfq_pk)
        item_fks_dict = {}


        # check for quote_fk_list
        if not quote_fk_list:
            LOGGER.error("MT API did not return anything for the last query.")
            raise ValueError("MT API did not return anything for the last query.")

        # check for quantity not being none
        qty_none_buf: List[int] = []
        for fk, qty in quote_fk_list:
            if not qty:
                qty_none_buf.append(fk)

        if qty_none_buf:
            LOGGER.error(f"QTY is None for the following fks: {qty_none_buf}")
            raise ValueError(f"Missing QTY for fks {qty_none_buf}")

        if quote_fk_list:
            for fk, quantity in quote_fk_list:
                # Retrieve the list of (ItemFK, QuantityRequired) for each QuoteFK
                item_pks = quote_assembly_table.get("ItemFK", "QuantityRequired", QuoteFK=fk)
                if item_pks:
                    for item_fk, qty_required in item_pks:
                        if qty_required is not None:
                            total_quantity = qty_required * quantity  # calculate the total for one line item
                            if item_fk in item_fks_dict:
                                item_fks_dict[item_fk] += total_quantity
                            else:
                                item_fks_dict[item_fk] = total_quantity

        item_and_details_dict = get_item_dict(item_fks_dict)
        for key, _ in item_and_details_dict.items():
            try:
                commodity = self.get_commodity_from_item(key)
            except ValueError as e:
                commodity = None
            item_and_details_dict[key]["commodity"] = commodity
        
        return item_and_details_dict


    def get_emailgroup_for_code(self, comm_code) -> List[str]:
        """
        TO BE USED WITH GET EMAIL FROM PARTY ONLY 
        Returns PartyPKs associated with the Code.
        Code can be found in Mie Trak Party Group Maintanence.

        :param code str: eg: "MAT-AL-PLT" "OP-FIN" etc
        :return: List[(pk,),...]
        :raises ValueError: If the code provided does not correspond to a party group.
        :raises ValueError: No PartyPKs are returned i.e. group is created but it does not have any parties in it.
        """
        customer_group_table = TableManger("CustomerGroup")
        customer_group_pk = customer_group_table.get("CustomerGroupPK", Code=comm_code)

        if not customer_group_pk:
            raise ValueError(f"Code: {comm_code} does not have a party group.")

        party_customer_group = TableManger("PartyCustomerGroup")
        party_fks = party_customer_group.get("PartyFK", CustomerGroupFK=customer_group_pk[0][0])  # multiple partyfks returned

        if not party_fks:
            raise ValueError(f"No parties in Party Group for code: {comm_code}")

        return self._get_email_from_party_pks(party_fks)


    def _get_email_from_party_pks(self, party_pks) -> List[str]:
        """
        Gets email IDs from Party Table when given a list of party Pks.

        :param party_pks: List of Party PKs to get the emails for. Must be in List[(<pk>, ), ...] format.
        :return: List of emails.
        """

        party_table = TableManger("Party")
        emails = []

        for party_pk in party_pks:
            email = party_table.get("Email", PartyPK=party_pk[0])
            emails.append(email[0][0])

        return emails

    def get_all_codes_and_emails(self) -> Dict[str, List[str]]:
        code_to_email_dict = {}
        commodity_table = TableManger("Commodity")
        codes = commodity_table.get("Code")
        for code in codes:
            try:
                emails = self.get_emailgroup_for_code(code[0])
            except ValueError:
                # TODO: log this error
                emails = ["No Emails Found"]

            code_to_email_dict[code[0]] = emails

        return code_to_email_dict

