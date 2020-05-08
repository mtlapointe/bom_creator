import argparse
import bomloader
import dfexporter
import misysloader
import easygui

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
full_bom_df = bomloader.BOM().load_csv(csv_file_path)

# Load MIsys PO Data
misys_po_df = misysloader.MisysTable().load_po_data()

# Create Excel Exporter
excel_export = dfexporter.DFExport(f'{full_bom_df.get_date_from_file()} {full_bom_df.get_assy_from_file()}.xlsx')

# Generate Full BOM Sheet

full_bom_cols = ['Level',
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

# 1 - Assembly Tree

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

excel_export.add_sheet(full_bom_df.df,
                       sheet_name='Assembly BOM',
                       freeze_col=5, freeze_row=1,
                       cols_to_print=assy_bom_cols,
                       depth_col_name='Depth',
                       group_rows=True,
                       highlight_depth=True,
                       highlight_col_limit=0,
                       cols_to_indent=['Part Number'],
                       print_index=True)

# Generate Purchasing Status List

shipset_qty = 1
part_bom_df['Shipset Qty Required'] = part_bom_df['Total QTY'] * shipset_qty

po_data_by_part_df = misys_po_df.groupby(['Product Number']) \
    .agg({'Qty Ordered': 'sum',
          'Qty Recd': 'sum',
          'PO Number': lambda x: ','.join(sorted(set(x))),
          'Promised Date': 'max'})

purch_list_df = part_bom_df.merge(po_data_by_part_df, how='left', left_on='Part Number', right_index=True) \
    .reset_index(drop=True)

purch_list_df.rename(columns={'Total QTY': 'Assy Qty Required', 'Promised Date': 'Next Recv Date'}, inplace=True)

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

excel_export.add_sheet(purch_list_df,
                       sheet_name='Purchasing BOM',
                       cols_to_print=purch_cols,
                       print_index=False)


# Create sheet with full PO data for parts in BOM

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

bom_po_data = misys_po_df.loc[misys_po_df['Product Number'].isin(full_bom_df.df['Part Number'])]

excel_export.add_sheet(bom_po_data,
                       sheet_name='PO Data',
                       cols_to_print=po_data_cols,
                       print_index=False)

excel_export.write_book()