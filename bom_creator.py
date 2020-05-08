""" BOM Creator Module

This module will load a BOM file from PDM (CSV or JSON) and create an Excel BOM tool.

"""

import argparse
import bomloader
import dfexporter
import misysloader
import easygui
import pandas as pd
import numpy as np
from datetime import datetime as dt
import odooloader


class BOMCreator:

    def __init__(self, export_file_name=None, csv_file_path=None):
        """ Initialize a BOMCreator object


        Args:
            export_file_name (str, optional): File name or full path for exported Excel file.
                If none given, the filename for the CSV will be used with a date stamp.
            csv_file_path (str, optional): File path for BOM CSV file to load.
                Required if using class directly without Drag-n-Drop batch file
        """

        # If using Drag and Drop - get CSV file name from arguments
        parser = argparse.ArgumentParser(description='Process CSV file')
        parser.add_argument('--file', default='')
        args, unknown = parser.parse_known_args()
        file_name = args.file

        # Check if file passed was .csv, if not set file path to None
        if '.csv' in file_name.lower():
            csv_file_path = file_name

        # Load BOM from CSV into DataFrame
        self.full_bom_df = bomloader.BOM().load_csv(csv_file_path)

        # Load MIsys PO Data - OBSOLETE
        # self.misys_po_df = misysloader.MisysTable(cache_age_limit=72).load_po_data()

        # Load Odoo PO Data in DataFrame.
        # Filter only items that are purchased (no RFQ's or cancelled orders)
        odoo_df = odooloader.OdooLoader().get_po_lines_df()
        self.odoo_po_df = odoo_df.loc[odoo_df['Status'] == 'purchase']

        # Create Excel DFExporter object
        if export_file_name is None:
            export_file_name = f'{self.full_bom_df.get_date_from_file()} {self.full_bom_df.get_assy_from_file()}.xlsx'
        self.excel_export = dfexporter.DFExport(export_file_name)

        # Create a DF with BOM info grouped by Part Number
        self.part_bom_df = self.create_part_bom_df()

        # Generate Full BOM Sheet
        self.full_bom_cols = ['Level',
                              'Depth',
                              'Type',
                              'Part Number',
                              'Description',
                              'QTY',
                              'Total QTY',
                              'Used On',
                              'Cage Code',
                              'Revision',
                              'Drawing',
                              'Drawing Number',
                              'Duplicate',
                              'Material',
                              'Finish 1',
                              'Finish 2',
                              'Finish 3',
                              'Weight',
                              'State',
                              'Latest Version',
                              'Unique ID',
                              'Parent ID',
                              'Parent List']

    def add_assy_bom_sheet(self):

        assy_bom_cols = ['Level',
                         'Depth',
                         'Type',
                         'Part Number',
                         'Description',
                         'QTY',
                         'Total QTY',
                         'Material',
                         'Weight',
                         'Used On',
                         'Cage Code',
                         'Revision',
                         'State',
                         'Latest Version',
                         'Drawing Number',
                         'Duplicate']

        self.excel_export.add_sheet(self.full_bom_df.df,
                                    sheet_name='Assembly BOM',
                                    freeze_col=5, freeze_row=1,
                                    cols_to_print=assy_bom_cols,
                                    depth_col_name='Depth',
                                    group_rows=True,
                                    highlight_depth=True,
                                    highlight_col_limit=0,
                                    cols_to_indent=['Part Number'],
                                    print_index=True)

    def add_drawing_list_sheet(self):

        drawing_bom_df = self.full_bom_df.df[self.full_bom_df.df['Drawing Number'].notnull()] \
            .drop_duplicates('Drawing Number') \
            .sort_values('Drawing Number')

        drw_bom_cols = ['Drawing Number',
                        'Description',
                        'State',
                        'Revision',
                        'Latest Version']

        self.excel_export.add_sheet(drawing_bom_df,
                                    sheet_name='Drawings',
                                    cols_to_print=drw_bom_cols,
                                    print_index=False)

    def add_drawing_tree(self):

        drawing_bom_df = self.full_bom_df.df[self.full_bom_df.df['Drawing Number'].notnull()] \
            .drop_duplicates('Drawing Number') \
            .sort_values('Drawing Number')

        drw_bom_cols = ['Drawing Number',
                        'Description',
                        'State',
                        'Revision',
                        'Latest Version']

        self.excel_export.add_sheet(drawing_bom_df,
                                    sheet_name='Drawings',
                                    cols_to_print=drw_bom_cols,
                                    print_index=False)

    def create_part_bom_df(self):
        """ Create a DataFrame from full BOM data that groups by Part Number

        Returns:
            DataFrame: DF with BOM data grouped by PN

        """

        part_bom_df = self.full_bom_df.df.groupby(['Part Number']).agg({'Total QTY': 'sum',
                                                                        'Description': 'first',
                                                                        'Cage Code': 'first',
                                                                        'Revision': 'first',
                                                                        'Material': 'first',
                                                                        'Finish 1': 'first',
                                                                        'Finish 2': 'first',
                                                                        'Finish 3': 'first',
                                                                        'Weight': 'first'
                                                                        }).reset_index()

        return part_bom_df['Weight'].fillna(0, inplace=True)

    def add_mnp_sheet(self):

        self.mnp_df = self.part_bom_df
        self.mnp_df['Weight'] = (pd.to_numeric(self.mnp_df['Weight'], errors='coerce').fillna(0))
        self.mnp_df['Total Weight'] = self.mnp_df['Total QTY'] * self.mnp_df['Weight']

        mnp_cols = ['Part Number',
                    'Revision',
                    'Description',
                    'Material',
                    'Finish 1',
                    'Finish 2',
                    'Finish 3',
                    'Total QTY',
                    'Weight',
                    'Total Weight']

        self.excel_export.add_sheet(self.mnp_df,
                                    sheet_name='M&P and Mass List',
                                    cols_to_print=mnp_cols,
                                    print_index=False)

    def add_purchasing_status_sheet(self, shipset_qty=None):

        if shipset_qty is None:
            shipset_qty = easygui.integerbox('How many shipsets to evaluate for purchasing?')

        self.part_bom_df['Shipset Qty Required'] = self.part_bom_df['Total QTY'] * shipset_qty

        # Merge the MISys and Odoo DF's. Drop NA's from the Due Date and then need to convert string to
        # DT because the merged DF shows them as strings.
        merged_po_df = self.misys_po_df.rename(columns={'Promised Date': 'Due Date'}) \
            .append(self.odoo_po_df, sort=True).dropna(subset=['Due Date'])
        merged_po_df['Due Date'] = pd.to_datetime(merged_po_df['Due Date'])

        po_data_by_part_df = merged_po_df.groupby(['Product Number']) \
            .agg({'Qty Ordered': 'sum',
                  'Qty Recd': 'sum',
                  'PO Number': lambda x: ', '.join(sorted(set(x))),
                  'Due Date': 'max'})

        purch_list_df = self.part_bom_df.merge(po_data_by_part_df, how='left', left_on='Part Number', right_index=True) \
            .reset_index(drop=True)

        purch_list_df.rename(columns={'Total QTY': 'Assy Qty Required', 'Due Date': 'Next Recv Date'},
                             inplace=True)

        purch_list_df.fillna({'Qty Required': 0, 'Qty Ordered': 0, 'Qty Recd': 0}, inplace=True)

        purch_cols = ['Part Number',
                      'Revision',
                      'Description',
                      'Cage Code',
                      'Assy Qty Required',
                      'Shipset Qty Required',
                      'Qty Ordered',
                      'Qty Recd',
                      'Next Recv Date',
                      'PO Number']

        purch_list_df['Qty Ordered'] = purch_list_df['Qty Ordered'].astype('int64')
        purch_list_df['Qty Recd'] = purch_list_df['Qty Recd'].astype('int64')

        self.excel_export.add_sheet(purch_list_df,
                                    sheet_name='Purchasing BOM',
                                    cols_to_print=purch_cols,
                                    print_index=False)

    def add_po_data_sheet(self):

        po_data_cols = ['PO Number',
                        'Supplier',
                        'PO Line Number',
                        'Status',
                        'Job ID',
                        'Product Number',
                        'Product Revision',
                        'Description',
                        'Comment',
                        'Qty Ordered',
                        'Qty Recd',
                        'UOM',
                        'UOM Conversion',
                        'Unit Price',
                        'Initial Due Date',
                        'Actual Due Date',
                        'Promised Date',
                        'Date Last Recd',
                        'Data Type',
                        'Location ID']

        bom_po_data = self.misys_po_df.loc[self.misys_po_df['Product Number'].isin(self.full_bom_df.df['Part Number'])]

        self.excel_export.add_sheet(bom_po_data,
                                    sheet_name='MISys PO Data',
                                    cols_to_print=po_data_cols,
                                    print_index=False)

    def add_odoo_po_data_sheet(self):

        odoo_po_cols = [
            'PO Number',
            'Supplier',
            'PO Line Number',
            'Status',
            'Job ID',
            'Product Number',
            'Product Revision',
            'Description',
            'Qty Ordered',
            'Qty Recd',
            'Due Date',
            'Unit Price',
            'Tax Price',
            'Total Price'
        ]

        bom_odoo_po_data = self.odoo_po_df.loc[self.odoo_po_df['Product Number'].isin(
            self.full_bom_df.df['Part Number'])]

        self.excel_export.add_sheet(bom_odoo_po_data,
                                    sheet_name='Odoo PO Data',
                                    cols_to_print=odoo_po_cols,
                                    print_index=False)

    def get_schedule_df(self):
        schedule_bom_cols = ['Level',
                             'Unique ID',
                             'Parent ID',
                             'Parent Index',
                             'Depth',
                             'Type',
                             'Part Number',
                             'Description',
                             'QTY',
                             'Total QTY',
                             'Used On',
                             'Cage Code',
                             'Revision',
                             'State',
                             'Latest Version',
                             'Drawing Number',
                             'Duplicate',
                             'Parent List',
                             'Lead Time',
                             'Start Date',
                             'Finish Date']

        # Setup custom lead time sheet and DF
        lead_time_sheet_name = 'Schedule Lead Times'

        lead_time_df = pd.DataFrame(data=[('DSS PART', 10, np.NaN), ('DSS ASSY', 2, np.NaN), ('COTS', 2, np.NaN)],
                                    columns=['Part Number', 'Lead Time', 'Finish Date'])
        lead_time_df['Lead Time'] = lead_time_df['Lead Time'].astype(pd.Int64Dtype())

        df = self.full_bom_df.df

        # Look-up Unique ID for given Level
        def lookup_parent_index(x):
            return int(df.loc[df['Unique ID'] == x].iloc[0].name) if (x and x is not np.nan) else 0

        df['Parent Index'] = df['Parent ID'].apply(lookup_parent_index).astype('Int64')

        # assy_groups = df.groupby(by='Parent Index')
        #
        # grouped_df = pd.DataFrame()
        #
        # for parent_index, group_df in assy_groups:
        #     grouped_df = grouped_df.append(df.loc[[parent_index]], sort=True)
        #     grouped_df = grouped_df.append(group_df.loc[group_df['Type'] != 'DSS ASSY'], sort=True)
        #
        # grouped_df.reset_index(drop=True, inplace=True)
        grouped_df = df

        # Get the column number for 'Type'
        type_col_num = 'MATCH("Type",$1:$1,0)'
        # Lookup the 'Type' for that row
        row_type_value = f'INDEX($1:$100000,ROW(),{type_col_num})'

        # Get the column number for 'Part Number'
        pn_col_num = 'MATCH("Part Number",$1:$1,0)'
        # Lookup the 'Part Number' for that row
        row_pn_value = f'INDEX($1:$100000,ROW(),{pn_col_num})'

        # Get the column numbers for headers on the Lead Time sheet
        lt_sheet_pn_col_num = f'MATCH("Part Number",\'{lead_time_sheet_name}\'!$1:$1,0)'
        lt_sheet_lead_time_col_num = f'MATCH("Lead Time",\'{lead_time_sheet_name}\'!$1:$1,0)'
        lt_sheet_finish_date_col_num = f'MATCH("Finish Date",\'{lead_time_sheet_name}\'!$1:$1,0)'

        # Get Lead Time sheet row number for given Part Number - returns error if none.
        # Then get the Lead Time and Finish Date values for that PN
        lt_sheet_pn_row = f'MATCH({row_pn_value},INDEX(\'{lead_time_sheet_name}\'!$1:$100000,,{lt_sheet_pn_col_num}),0)'
        lt_sheet_pn_lead_time_val = f'INDEX(\'{lead_time_sheet_name}\'!$1:$100000,' \
                                    f'{lt_sheet_pn_row},{lt_sheet_lead_time_col_num})'
        lt_sheet_pn_finish_date_val = f'INDEX(\'{lead_time_sheet_name}\'!$1:$100000,' \
                                      f'{lt_sheet_pn_row},{lt_sheet_finish_date_col_num})'

        # Get default Lead Time for Type
        lt_sheet_type_row = f'MATCH({row_type_value},' \
                            f'INDEX(\'{lead_time_sheet_name}\'!$1:$100000,,{lt_sheet_pn_col_num}),0)'
        lt_sheet_type_lead_time_val = f'INDEX(\'{lead_time_sheet_name}\'!$1:$100000,' \
                                      f'{lt_sheet_type_row},{lt_sheet_lead_time_col_num})'

        # lead_time_equation = f'_xlfn.SWITCH({row_type_value},"DSS ASSY",2,"DSS PART",10,"COTS",2)'
        # If the lookup for PN in the Lead Time sheet is error or blank, return default Type Lead-time
        lead_time_equation = f'IF(' \
                             f'OR(ISERROR({lt_sheet_pn_lead_time_val}),ISBLANK({lt_sheet_pn_lead_time_val})),' \
                             f'{lt_sheet_type_lead_time_val},' \
                             f'{lt_sheet_pn_lead_time_val}' \
                             f')'

        # Get column positions
        parent_id_col_num = 'MATCH("Parent ID",$1:$1,0)'
        unique_id_col_num = 'MATCH("Unique ID",$1:$1,0)'
        start_col_num = 'MATCH("Start Date",$1:$1,0)'
        finish_col_num = 'MATCH("Finish Date",$1:$1,0)'
        lead_time_col_num = 'MATCH("Lead Time",$1:$1,0)'

        # Get the Parent ID for the row
        row_parent_id_value = f'INDEX($1:$100000,ROW(),{parent_id_col_num})'
        # Lookup the row number of that Parent ID
        row_parent_id_row_num = f'MATCH({row_parent_id_value},INDEX($1:$100000,,{unique_id_col_num}),0)'
        # Lookup the Start Date of the Parent
        finish_date_equation = f'WORKDAY(INDEX($1:$100000,{row_parent_id_row_num},{start_col_num}),-1)'

        # Get the Lead Time for the row
        row_lead_time_value = f'INDEX($1:$100000,ROW(),{lead_time_col_num})'
        # Subtract Lead Time from Finish Date
        start_date_equation = f'=WORKDAY(INDEX($1:$100000,ROW(),{finish_col_num}),-{row_lead_time_value}*5)'

        grouped_df['Lead Time'] = f'={lead_time_equation}'
        # grouped_df['Finish Date'] = f'=IFERROR({lt_sheet_pn_finish_date_val},{finish_date_equation}'
        grouped_df['Finish Date'] = f'=IF(' \
                                    f'OR(ISERROR({lt_sheet_pn_finish_date_val}),' \
                                    f'ISBLANK({lt_sheet_pn_finish_date_val})),' \
                                    f'{finish_date_equation},' \
                                    f'{lt_sheet_pn_finish_date_val}' \
                                    f')'
        grouped_df['Start Date'] = start_date_equation

        # Set a default lead-time for the top-item
        lead_time_df = lead_time_df.append({
            'Part Number': str(grouped_df.loc[0, 'Part Number']), 'Finish Date': 100000}, ignore_index=True)

        # Add the lead-time sheet
        self.excel_export.add_sheet(lead_time_df,
                                    sheet_name=lead_time_sheet_name,
                                    print_index=False)

        return grouped_df[schedule_bom_cols]

    def add_schedule_bom_sheet(self):

        df = self.get_schedule_df()

        date_format = {'num_format': 'mm/dd/yy',
                       'bold': False,
                       'border': 1}

        self.excel_export.add_sheet(df,
                                    sheet_name='Schedule BOM',
                                    freeze_col=7, freeze_row=1,
                                    print_index=True,
                                    col_formats={'Start': 'custom', 'Finish': 'custom', 'Cage Code': 'string',
                                                 'Level': 'string'},
                                    col_style={'Start': date_format, 'Finish': date_format})

    def add_debug_sheet(self):
        self.excel_export.add_raw_sheet(self.full_bom_df.df, 'Debug')

    def write_book(self):
        self.excel_export.write_book()


def main():
    bom_creator = BOMCreator()
    bom_creator.add_assy_bom_sheet()
    bom_creator.add_schedule_bom_sheet()
    bom_creator.add_drawing_list_sheet()
    bom_creator.add_mnp_sheet()
    bom_creator.add_purchasing_status_sheet()
    bom_creator.add_po_data_sheet()
    bom_creator.add_odoo_po_data_sheet()
    bom_creator.write_book()


def main_test():
    global bom_creator

    test_file_name = f'test {int(dt.now().timestamp())}.xlsx'
    bom_creator = BOMCreator(export_file_name=test_file_name, csv_file_path='test.csv')
    bom_creator.add_assy_bom_sheet()
    bom_creator.add_schedule_bom_sheet()
    bom_creator.add_drawing_list_sheet()
    bom_creator.add_mnp_sheet()
    bom_creator.add_purchasing_status_sheet(shipset_qty=2)
    bom_creator.add_po_data_sheet()
    bom_creator.add_odoo_po_data_sheet()
    bom_creator.write_book()


def test_load():
    global bom_creator, full_df, sched_df

    bom_creator = BOMCreator(csv_file_path='test.csv')
    full_df = bom_creator.full_bom_df.df
    sched_df = bom_creator.get_schedule_df()


if __name__ == '__main__':
    main()

if __name__ == 'builtins':
    # test_load()
    main_test()
