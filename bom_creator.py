import argparse
import bomloader
import dfexporter

# If using Drag and Drop - get CSV file name from arguments
parser = argparse.ArgumentParser(description='Process CSV file')
parser.add_argument('--file', default='')
args, unknown = parser.parse_known_args()
file_name = args.file

# Check if file passes was .csv, if not set file path to None
csv_file_path = None
if '.csv' in file_name:
    csv_file_path = file_name

bom = bomloader.BOM().load_csv(csv_file_path)

cols = ['Level', 'Depth', 'Type', 'Part Number', 'Description', 'QTY', 'Total QTY', 'Used On', 'Cage Code',
        'Revision', 'Drawing', 'Drawing Number', 'Duplicate', 'Material', 'Finish 1', 'Finish 2', 'Finish 3',
        'Weight', 'State', 'Latest Version']

export = dfexporter.DFExport(f'{bom.get_date_from_file()} {bom.get_assy_from_file()}.xlsx')
export.add_sheet(bom.df,
                 sheet_name='Highlight No Index',
                 freeze_col=10, freeze_row=1,
                 cols_to_print=cols,
                 depth_col_name='Depth', group_rows=True,
                 highlight_rows=True, highlight_col_limit=0,
                 cols_to_indent=['Part Number'],
                 print_index=True)

export.write_book()



