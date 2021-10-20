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
    def __init__(self, file_path, header_row_index, columns_to_keep):
        super().__init__(file_path, header_row_index, columns_to_keep)
        

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
                    if row_values[0] == 'Transaction Fee': # End of workbook check
                        return
                    self.output_table.append(row_values)
        return
        

class PriceList(POSReport):
    def __init__(self, file_path, header_row_index, columns_to_keep):
        super().__init__(file_path, header_row_index, columns_to_keep)

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
                    if row_values[0] == 'Transaction Fee': # End of workbook check
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


class CleanedItemList:
    def __init__(self, active_upc_item_file_path):
        self.active_upc_item_file_path = active_upc_item_file_path

    def extract_unit_size(self):
        pass


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
    pd.DataFrame(filtered_list).to_excel('items_with_unit_size.xls')


if __name__ == "__main__":
    all_sales_items_combined = []
    # Read in and append SalesList objects to the list
    for i in range(4,10):
        all_sales_items_combined.append(SalesList(f'./pos_reports/sales/sales_{i}.xls', 8, 2))
    
    # Filter each SalesList item and convert to a DataFrame
    for i, sales_list in enumerate(all_sales_items_combined):
        sales_list.filter_rows()
        all_sales_items_combined[i] = pd.DataFrame(sales_list.output_table, columns=['UPC', 'Item Name'])
    
    all_sales_items_combined_df = pd.concat(all_sales_items_combined)
    all_active_items = all_sales_items_combined_df.drop_duplicates()
    all_active_items.to_excel('active_items_past_6_months.xls')

    """duplicate_items = all_active_items[all_active_items['Item Name'].duplicated() == True]
    duplicate_items.to_excel('duplicate_items.xls')

    duplicate_upc = all_active_items[all_active_items['UPC'].duplicated() == True]
    duplicate_upc.to_excel('duplicate_upc.xls')"""
    all_active_items_list = list(all_active_items['Item Name'].astype(str))
    match_expression_on_list(all_active_items_list)

    """#print(len(no_duplicates))
    inventory_list = InventoryList('./pos_reports/Inventory List_2021_10_14.xls', 5, 3)
    inventory_list.filter_rows()
    print('inventory', len(inventory_list.output_table))
    inventory_list.remove_duplicate_rows()
    print('inventory', len(inventory_list.output_table))

    price_list = PriceList('./pos_reports/Price List_2021_10_14.xls', 4, 3)
    price_list.filter_rows()
    print('price list', len(price_list.output_table))
    price_list.remove_duplicate_rows()
    print('price list', len(price_list.output_table))"""

