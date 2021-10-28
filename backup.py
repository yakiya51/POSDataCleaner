import xlrd
import pandas as pd
import re

class POSReport:
    def __init__(self, file_path: str, header_row_index: int, columns_to_keep: list[str]):
        self.file_path = file_path
        self.header_row_index = header_row_index
        self.columns_to_keep = columns_to_keep
        self.output_table = []

        workbook = xlrd.open_workbook(filename=self.file_path)
        worksheet = workbook.sheet_by_index(0)
        # Generator object
        self.row_generator = worksheet.get_rows()

        # Extract Header Row (Column Names)
        self.header_row = next([x.value for x in row[:self.columns_to_keep]] for i,row in enumerate(self.row_generator) if i==header_row_index-1)

    def remove_duplicate_rows(self):
        """Removes Duplicate rows within the output rows"""
        self.output_table = set(tuple(row) for row in self.output_table)


class InventoryList(POSReport):
    def __init__(self, file_path, header_row_index=5, columns_to_keep=3):
        super().__init__(file_path, header_row_index, columns_to_keep)
        self.filter_rows()

    def filter_rows(self):
        """
        Filter Steps:
        - Disregard contents in columns D onwards
        - Remove blank rows
        - Remove irrelevat rows (non-item rows) such as column, category, total rows
        - Remove Duplicate Rows
        """
        for row in self.row_generator:
            row_values = [x.value for x in row[:self.columns_to_keep]]
            if any(row_values): # Empty Row check
                if row_values[1] and row_values[2]: # Valid Item check
                    if row_values[0] == 'Transaction Fee' or row_values[1] == 'Transaction Fee': # End of workbook check
                        return
                    self.output_table.append(row_values)
        return
        

class PriceList(POSReport):
    def __init__(self, file_path, header_row_index=4, columns_to_keep=3):
        super().__init__(file_path, header_row_index, columns_to_keep)
        self.filter_rows()

    def filter_rows(self):
        """
        Filter Steps:
        - Disregard contents in columns D through G
        - Remove blank rows
        - Remove irrelevat rows (non-item rows) such as column, category, total rows
        - Strip/Upper each item name & size
        - Try and extract the unit/packaging size
        - Remove Duplicate Rows
        """
        for row in self.row_generator:
            row_values = [x.value.upper() if isinstance(x.value, str) else x.value for x in row[:self.columns_to_keep]]
            if any(row_values): # Empty Row check
                if row_values[0] and row_values[2]: # Valid Item check
                    if row_values[0] == 'Transaction Fee' or row_values[1] == 'Transaction Fee': # End of workbook check
                        return
                    self.output_table.append(row_values)
        return
           

class SalesList(POSReport):
    def __init__(self, file_path, header_row_index, columns_to_keep):
        super().__init__(file_path, header_row_index, columns_to_keep)

    def filter_rows(self):
        """
        Extract only the items and their UPC on the sales list
        """
        for row in self.row_generator:
            row_values = [x.value.upper() if isinstance(x.value, str) else x.value for x in row[:self.columns_to_keep]]
            if any(row_values): # Empty Row check
                if row_values[0] and row_values[1]: # Valid Item check
                    if row_values[0] != 'UPC': 
                        self.output_table.append(row_values)
        return


class FullItemList:
    def __init__(self, active_upc_item_file_path):
        self.active_upc_item_file_path = active_upc_item_file_path
        self.df = pd.read_excel(active_upc_item_file_path)
        self.create_unit_size_column()
        self.df.to_excel('full_item_list.xls', index=False)

    def create_unit_size_column(self):
        self.df['Unit Size'] = self.df['Item Name'].apply(lambda x: self.extract_unit_size(x))
        self.df['Item Name'] = self.df['Item Name'].apply(lambda x: self.remove_unit_size(x))

        
        return

    def extract_unit_size(self, x):
        item_name = str(x)
        filter = re.compile('\d+\.*\d* *(ML|L|PAK|OZ|GALLON|CAN|OZ|L|P|O|BLT|BTL|PACK|CT|0Z|BTLS|OZ BOTTLE|CN|G|LBS|QT|Z|STICKS)+')
        if re.search(filter, item_name):
            return item_name[re.search(filter, item_name).start():re.search(filter, item_name).end()].strip()
        else:
            return ''

    def remove_unit_size(self, x):
        item_name = str(x)
        filter = re.compile('\d+\.*\d* *(ML|L|PAK|OZ|GALLON|CAN|OZ|L|P|O|BLT|BTL|PACK|CT|0Z|BTLS|OZ BOTTLE|CN|G|LBS|QT|Z)+')
        if re.search(filter, item_name):
            return (item_name[:re.search(filter, item_name).start()] + item_name[re.search(filter, item_name).end()+1:]).strip()
        else:
            return item_name
        
def match_expression_on_list(item_list):
    """
    - PAK
    - OZ
    - GALLON
    - CAN
    - OZ
    - L
    - P
    - O
    - ML
    - BLT
    - BTL
    - PACK
    - CT
    - 0Z
    - BTLS
    - OZ BOTTLE
    - CN
    - G
    - LBS
    - QT
    - Z
    """
    filter = re.compile('\d+\.*\d* *([A-Z])+')
    filtered_list = []
    for item in item_list:
        if re.findall(filter, item):
            filtered_list.append(item)
    pd.DataFrame(filtered_list).to_excel('items_with_unit_size.xls', index=False)


if __name__ == "__main__":
    inventory = InventoryList('./pos_reports/inventory/inv_10_14_2021.xls')
    price = PriceList('./pos_reports/price/price_10_14_2021.xls')
    print(price.output_table[:100])
    #inventory_price_list = [ i + j[1:] for i in inventory.output_table for j in price.output_table if i[1] == j[0] ]
    #print(inventory_price_list)



    """all_sales_items_combined = []
    # Read in and append SalesList objects to the list
    for i in range(4,10):
        all_sales_items_combined.append(SalesList(f'./pos_reports/sales/sales_{i}.xls', 8, 2))
    
    # Filter each SalesList item and convert to a DataFrame
    for i, sales_list in enumerate(all_sales_items_combined):
        sales_list.filter_rows()
        all_sales_items_combined[i] = pd.DataFrame(sales_list.output_table, columns=['UPC', 'Item Name'])
    
    all_sales_items_combined_df = pd.concat(all_sales_items_combined)
    all_active_items = all_sales_items_combined_df.drop_duplicates()
    all_active_items.to_excel('active_items_past_6_months.xls', index=False)

    duplicate_items = all_active_items[all_active_items['Item Name'].duplicated() == True]
    duplicate_items.to_excel('duplicate_items.xls', index=False)

    duplicate_upc = all_active_items[all_active_items['UPC'].duplicated() == True]
    duplicate_upc.to_excel('duplicate_upc.xls', index=False)
    all_active_items_list = list(all_active_items['Item Name'].astype(str))
    match_expression_on_list(all_active_items_list)

    FullItemList('active_items_past_6_months.xls')"""



