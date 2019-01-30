import argparse
import bomloader
import dfexporter
import misysloader
import easygui
import numpy as np
import datetime

# If using Drag and Drop - get CSV file name from arguments
parser = argparse.ArgumentParser(description='Process CSV file')
parser.add_argument('--file', default='')
args, unknown = parser.parse_known_args()
file_name = args.file

# Check if file passed was .csv, if not set file path to None
csv_file_path = None
if '.csv' in file_name.lower():
    csv_file_path = file_name

# Load BOM from CSV
bom = bomloader.BOM().load_csv(csv_file_path)

# Load MIsys PO Data
misys = misysloader.MisysTable().load_po_data()

# Create Excel Exporter
export = dfexporter.DFExport(f'{bom.get_date_from_file()} {bom.get_assy_from_file()}.xlsx')

# Generate Full BOM Sheet

full_bom_cols = ['Level', 'Depth', 'Type', 'Part Number', 'Description', 'QTY', 'Total QTY', 'Used On', 'Cage Code',
                 'Revision', 'Drawing', 'Drawing Number', 'Duplicate', 'Material', 'Finish 1', 'Finish 2', 'Finish 3',
                 'Weight', 'State', 'Latest Version']

assy_bom_cols =  ['Level', 'Depth', 'Type', 'Part Number', 'Description', 'QTY', 'Total QTY', 'Used On', 'Cage Code',
                  'Revision', 'State', 'Latest Version', 'Drawing Number', 'Duplicate']


export.add_sheet(bom.df,
                 sheet_name='Assembly BOM',
                 freeze_col=5, freeze_row=1,
                 cols_to_print=assy_bom_cols,
                 depth_col_name='Depth',
                 group_rows=True,
                 highlight_depth=True,
                 highlight_col_limit=0,
                 cols_to_indent=['Part Number'],
                 print_index=True)

# Generate Drawing List

drw_df = bom.df.drop_duplicates('Drawing Number').sort_values('Drawing Number')

drw_bom_cols = ['Drawing Number', 'Description', 'State', 'Revision', 'Latest Version']

export.add_sheet(drw_df,
                 sheet_name='Drawings',
                 cols_to_print=drw_bom_cols,
                 print_index=False)

# Generate Purchasing List

shipset_qty = easygui.integerbox('How many shipsets to evaluate for purchasing?')

purch_bom = bom.df.groupby(['Part Number']).agg({'Total QTY': 'sum',
                                                 'Description': 'first',
                                                 'Cage Code': 'first',
                                                 'Revision': 'first'})

purch_bom['Shipset Qty Required'] = purch_bom['Total QTY'] * shipset_qty

po_data = misys.groupby(['Product Number']).agg({'Qty Ordered': 'sum',
                                                 'Qty Recd': 'sum',
                                                 'PO Number': lambda x: ','.join(list(x)),
                                                 'Promised Date': 'max'})

purch_df = purch_bom.merge(po_data, how='left', left_index=True, right_index=True).reset_index()

purch_df.rename(columns={'Total QTY': 'Assy Qty Required', 'Promised Date':'Next Recv Date'}, inplace=True)

purch_df.fillna({'Qty Required':0, 'Qty Ordered':0, 'Qty Recd':0}, inplace=True)

purch_cols = ['Part Number', 'Revision', 'Description', 'Cage Code', 'Assy Qty Required', 'Shipset Qty Required',
              'Qty Ordered', 'Qty Recd', 'Next Recv Date', 'PO Number']

export.add_sheet(purch_df,
                 sheet_name='Purchasing BOM',
                 cols_to_print=purch_cols,
                 print_index=False)

# PO Data

po_data_cols = ['PO Number', 'Supplier', 'PO Line Number', 'Status', 'Job ID',
                'Product Number', 'Product Revision', 'Description', 'Comment',
                'Qty Ordered', 'Qty Recd', 'UOM', 'UOM Conversion', 'Unit Price',
                'Initial Due Date', 'Actual Due Date', 'Promised Date',
                'Date Last Recd', 'Data Type', 'Location ID']

bom_po_data = misys.loc[misys['Product Number'].isin(bom.df['Part Number'])]


export.add_sheet(bom_po_data,
                 sheet_name='PO Data',
                 cols_to_print=po_data_cols,
                 print_index=False)


# Print debug sheet


export.add_raw_sheet(drw_df, 'Debug')

export.write_book()
