import mt_commodity_script as cs
from commodity_bucket import CommodityBucket 
import helper
from pprint import pprint


test_rfq_pk = 5254
item_details_dict = helper.get_item_dict(helper.get_item_pks(test_rfq_pk))
pprint(item_details_dict)


# found_commodities = {}
# for key, value in item_pks_dict:
#     item_details_dict = cs.get_item_details(key)
#     item_commodity_code = cs.get_commodity_from_item(key)
#
#     if not item_commodity_code in found_commodities.keys():
#         found_commodities[item_commodity_code] = CommodityBucket(item_commodity_code)
#         found_commodities[item_commodity_code].items.append()

