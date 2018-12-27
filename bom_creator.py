import easygui
import pandas as pd
import numpy as np
import argparse

# If using Drag and Drop - get CSV file name from arguments
parser = argparse.ArgumentParser(description='Process CSV file')
parser.add_argument('--file', default='')
args, unknown = parser.parse_known_args()
file_name = args.file

# If File Name is not in Arguments, use GUI file selection
if '.csv' in file_name:
    bom_file = file_name
else:
    bom_file = easygui.fileopenbox(msg='Choose requirement spreadsheet', default='*.csv',
                                filetypes=[["*.csv", "All files"]])

# Load into DF
df = pd.read_csv(bom_file, encoding='utf_16', dtype={'Level': object}, float_precision='round_trip')

# Add level column if it doesn't exist
if not 'Level' in df.columns:
    df['Level'] = ''

# Fix filename - force uppercase and split off extension
df['File Name'], df['Extension'] = df['Name'].str.strip().str.upper().str.rsplit('.', n=1).str

# Drop any non-SW file from the list
df = df[df['Extension'].isin(['SLDPRT','SLDASM'])].reset_index(drop=True)

# Reset Part Number Column (crap data)
df['Part Number'] = np.NaN

# Check for NOCONFIG and assign File Name only to Part Number
df.loc[df['Configuration'].fillna('NOCONFIG').str.upper()=='NOCONFIG',
       ['Part Number']] = df['File Name']

# Check for PN Override and use that if true
df.loc[df['PartNumOverride'].notnull(), ['Part Number']] = df['PartNumOverride']

# Everything else, set PN to FILENAME + CONFIG
df.loc[df['Part Number'].isnull(), ['Part Number']] = df['File Name'] + df['Configuration']

# Determine type of part (DSS part/assy or COTS)
dss_part_filter = df['Part Number'].str.contains('[1,2][0-9]{2}[F,Q,N,G,E,X][0-9]{4}', na=False)
df.loc[dss_part_filter & (df['Extension']=='SLDPRT'),'Type'] = 'DSS PART'
df.loc[dss_part_filter & (df['Extension']=='SLDASM'),'Type'] = 'DSS ASSY'
df.loc[~dss_part_filter,'Type'] = 'COTS'

# Determine if drawing (DSS -1 number and no duplicate)
dash1_filter = df['Part Number'].str.contains('[1,2][0-9]{2}[F,Q,N,G,E,X][0-9]{4}-1', na=False)
df.loc[dash1_filter & (~df.duplicated('Part Number','first')),'Drawing'] = 'Yes'

# Mark duplicate parts
df.loc[df.duplicated('Part Number','first'), 'Duplicate'] = 'Yes'

# Determine depth of line (count level dots)
df['Depth'] = df.loc[~df['Level'].isnull(),'Level'].astype('str').str.count('[.]')

# Assign N/A for Material on Assemblies
df.loc[(df['Material'].isnull()) & (df['Extension']=='SLDASM'), 'Material'] = 'N/A - Assembly'

df['Total QTY'] = df['QTY']

for row_num, row in df.iterrows():
    row_depth = row['Depth']

    # Filter all rows after current with the same depth
    same_depth_rows = df.loc[row_num+1:].loc[df['Depth'] <= row_depth]

    # Get index of next row at same depth or set to None
    next_row = same_depth_rows.head().index[0] if len(same_depth_rows) else 0

    df.loc[row_num, 'Next Row'] = next_row

    # Multiply all lower rows by current row QTY
    if next_row > row_num + 1:
        df.loc[row_num + 1 : next_row - 1, 'Total QTY'] *= row['QTY']
        df.loc[row_num + 1: next_row - 1, 'Used On'] = row['Part Number']


parts_df_group = df.loc[df['Type'] != 'DSS ASSY'].groupby(['Part Number', 'Description'])
part_sum_df = parts_df_group['Total QTY'].agg(np.sum).reset_index()


# OUTPUT DATA TO EXCEL

output_df = df[['Level','Depth','Type','Part Number', 'Description', 'QTY', 'Total QTY', 'Used On', 'Cage Code',
       'Revision', 'Drawing', 'Duplicate', 'Material', 'Weight', 'State', 'Latest Version']].where((pd.notnull(df)), '')


# Setup Excel writer
writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter', options={'nan_inf_to_errors': True})

# Write data to Excel
output_df.to_excel(writer, sheet_name='BOM', startrow=1, startcol=1, header=False, index=False)

# Setup workbook and worksheet objects
workbook = writer.book
worksheet = writer.sheets['BOM']

# Set zoom and freeze panes location
worksheet.set_zoom(85)
worksheet.freeze_panes(1, 6)


# Setup styles
default_style = {
    'bold': False,
    'border': 1
}

header_style = {
    'rotation': 90,
    'align': 'Left',
    'bold': True
}

index_style = {
    'bold': True
}

# http://www.color-hex.com/color-palette/2280
depth_colors=[
    {'bg_color': '#3a4858',
     'font_color': 'white'},
    {'bg_color': '#757e8a',
     'font_color': 'white'},
    {'bg_color': '#b0b5bc',
     'font_color': 'black'},
    {'bg_color': '#ebecee',
     'font_color': 'black'},
    {'bg_color': 'white',
     'font_color': 'black'},
    {'bg_color': 'white',
     'font_color': 'black'},
    {'bg_color': 'white',
     'font_color': 'black'},
]


default_format = workbook.add_format({**default_style})
header_format = workbook.add_format({**default_style, **header_style})
index_format = workbook.add_format({**default_style, **index_style})

depth_format = []
for format in depth_colors:
    depth_format.append(workbook.add_format({**default_style, **format}))

indent_format = []
for i in range(0,10):
    indent_format.append(workbook.add_format({**default_style,**{'indent': i}}))
    if i < len(depth_colors):
        indent_format[i].set_bg_color(depth_colors[i]['bg_color'])
        indent_format[i].set_font_color(depth_colors[i]['font_color'])


# Write the column headers with the defined format.

worksheet.write(0, 0, 'Index', header_format)
for col_num, value in enumerate(output_df.columns.values):
    worksheet.write(0, col_num + 1, value, header_format)

# Write index column (no format)
for row_num, value in enumerate(output_df.index.values):
    worksheet.write(row_num + 1, 0, value, index_format)


for row_num, row in output_df.iterrows():
    depth = int(row['Depth'])

    # Write index column first
    worksheet.write(row_num + 1, 0, row_num, depth_format[depth])

    # Write rest of the row
    for col_num in range(len(row)):
        if output_df.dtypes[col_num] == 'int64':
            worksheet.write_number(row_num + 1, col_num + 1, row[col_num], default_format)
        else:
            worksheet.write_string(row_num + 1, col_num + 1, str(row[col_num]), default_format)


    # Write rest of the row
    for col_num in range(8):
        if output_df.dtypes[col_num] == 'int64':
            worksheet.write_number(row_num + 1, col_num + 1, row[col_num], depth_format[depth])
        else:
            worksheet.write_string(row_num + 1, col_num + 1, str(row[col_num]), depth_format[depth])

    # Re-write part number to get indent
    part_number = row['Part Number']
    worksheet.write(row_num + 1, 4, part_number, indent_format[depth])

    if depth > 0:
        worksheet.set_row(row_num + 1, None, None, {'level': depth})


def get_col_widths(dataframe):
    # First we find the maximum length of the index column
    idx_max = max([len(str(s)) for s in dataframe.index.values]) #+ [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    #return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values]) for col in dataframe.columns]

for i, width in enumerate(get_col_widths(output_df)):
    worksheet.set_column(i, i, width+2)


part_sum_df.to_excel(writer, sheet_name='Parts List')

writer.save()
