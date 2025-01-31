from mt_api.general_class import TableManger
from typing import List, Tuple, Dict


class CommodityBucket:
    def __init__(self, commodity_type) -> None:
        self.comm_code = commodity_type
        self.items = []
        self.get_email_from_party_pks()

    def _get_party_pks_from_code(self) -> List[Tuple]:
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
        customer_group_pk = customer_group_table.get("CustomerGroupPK", Code=self.comm_code)

        if not customer_group_pk:
            raise ValueError(f"Code: {self.comm_code} does not have a party group.")

        party_customer_group = TableManger("PartyCustomerGroup")
        party_fks = party_customer_group.get("PartyFK", CustomerGroupFK=customer_group_pk[0][0])  # multiple partyfks returned

        if not party_fks:
            raise ValueError(f"No parties in Party Group for code: {self.comm_code}")

        return party_fks

    def get_email_from_party_pks(self) -> List[str]:
        """
        Gets email IDs from Party Table when given a list of party Pks.

        :param party_pks: List of Party PKs to get the emails for. Must be in List[(<pk>, ), ...] format.
        :return: List of emails.
        """
        party_pks = self._get_party_pks_from_code()

        party_table = TableManger("Party")
        emails = []

        for party_pk in party_pks:
            email = party_table.get("Email", PartyPK=party_pk[0])
            emails.append(email[0][0])

        return emails
