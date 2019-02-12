import pandas as pd
from datetime import datetime as dt
import xlsxwriter

class DFExport:

    def __init__(self, output_file_name = "output.xlsx"):
        """ Create BOMExporter object. Use given output file name or default. """

        self.output_file_name = output_file_name

        self.workbook = xlsxwriter.Workbook(self.output_file_name, {'nan_inf_to_errors': True,
                                                                    'default_date_format': 'dd/mm/yy',
                                                                    'strings_to_numbers': True})

        self.__load_default_style()


    def __load_default_style(self):
        """ Setup default sheet styles for xlsxwriter """

        base_styles={}

        base_styles['default'] = {
            'bold': False,
            'border': 1
        }

        base_styles['index'] = {
            **base_styles['default'],
            'bold': True,
            'align': 'Center',
        }

        base_styles['float'] = {
            **base_styles['default'],
            'num_format': '0.00000'
        }

        base_styles['date'] = {
            **base_styles['default'],
            'num_format': 'mm/dd/yy'
        }

        header_style = {
            **base_styles['default'],
            'rotation': 90,
            'align': 'Center',
            'bold': True
        }
        self.header_format = self.workbook.add_format(header_style)

        # https://www.ibm.com/design/language/resources/color-library/
        depth_colors = [
            {'bg_color': '#464646',
             'font_color': 'white'},
            {'bg_color': '#595859',
             'font_color': 'white'},
            {'bg_color': '#777677',
             'font_color': 'white'},
            {'bg_color': '#949394',
             'font_color': 'black'},
            {'bg_color': '#a6a5a6',
             'font_color': 'black'},
            {'bg_color': '#c0bfc0',
             'font_color': 'black'},
            {'bg_color': '#d8d8d8',
             'font_color': 'black'},
            {'bg_color': '#eaeaea',
             'font_color': 'black'},
            {'bg_color': 'white',
             'font_color': 'black'},
        ]


        cell_styles = {}
        for base, base_style in base_styles.items():

            cell_styles[(base,None,None)] = base_style

            for depth in range(10):
                # Get depth colors, or use default if out of range
                colors = depth_colors[depth] if depth<len(depth_colors) else {}

                highlight_style = {**base_style, **colors}
                indent_style = {**base_style, 'indent': depth}

                cell_styles[(base,depth,'highlight')] = highlight_style
                cell_styles[(base,depth,'indent')] = indent_style
                cell_styles[(base,depth,'indent_highlight')] = {**indent_style, **highlight_style}

        self.cell_format = {}
        for cell_type, style in cell_styles.items():
            self.cell_format[cell_type] = self.workbook.add_format(style)



    def add_sheet(self, df, sheet_name="Sheet1", zoom=85, freeze_row=1, freeze_col=0, cols_to_print=None,
                  depth_col_name='', cols_to_indent=None, highlight_depth=False, highlight_col_limit=0, group_rows=False,
                  print_index=True):
        """ Take DF and creates new sheet with various options. """

        # Create output DF with only cols to print and replace N/A with empty string
        if cols_to_print:
            output_df = df[cols_to_print] #.where((pd.notnull(df)), '')
        else:
            output_df = df #.where((pd.notnull(df)), '')

        # If index column exists, need offset to shift all other columns
        index_col_offset = 1 if print_index else 0

        # Write data to Excel

        worksheet = self.workbook.add_worksheet(sheet_name)

        # Set zoom and freeze panes location
        worksheet.set_zoom(zoom)
        worksheet.freeze_panes(freeze_row, freeze_col)

        # Write the column headers with the defined format.
        if print_index:
            worksheet.write(0, 0, 'Index', self.header_format)
        for col_num, value in enumerate(output_df.columns.values):
            worksheet.write(0, col_num + index_col_offset, value, self.header_format)

        # Iterate through DF rows and write to Excel file
        for row_num in range(len(output_df)):

            # Get the row depth (if needed for highlight, indent or grouping)
            if highlight_depth or cols_to_indent or group_rows:
                depth = int(df[depth_col_name].iloc[row_num])
            else:
                depth = None

            format_option = 'highlight' if highlight_depth else None

            # Write optional index first using highlighted or plain index format
            print_format = self.cell_format[('index', depth, format_option)]
            if print_index:
                worksheet.write(row_num + 1, 0, output_df.index[row_num], print_format)

            # Write rest of the row
            for col_num in range(len(output_df.columns)):

                # Check if column should be highlighted and/or indented
                indent_col = cols_to_indent is not None and output_df.columns[col_num] in cols_to_indent
                highlight_col = highlight_depth and \
                                (highlight_col_limit==0 or col_num < highlight_col_limit-index_col_offset)

                # Choose the correct format option to use
                if indent_col and highlight_col:
                    format_option = 'indent_highlight'
                elif indent_col:
                    format_option = 'indent'
                elif highlight_col:
                    format_option = 'highlight'
                else:
                    format_option = None

                # Get value from DF
                df_value = output_df.iloc[row_num, col_num]

                # Set as empty string if null
                value = df_value if pd.notnull(df_value) else ''
                value_type = output_df.dtypes[col_num] if pd.notnull(df_value) else None

                # Write data as number or string
                if value_type in ['float64']:
                    worksheet.write_number(row_num + 1, col_num + index_col_offset, value,
                                           self.cell_format[('float', depth, format_option)])
                elif value_type in ['int64']:
                    worksheet.write_number(row_num + 1, col_num + index_col_offset, value,
                                           self.cell_format[('default', depth, format_option)])

                elif value_type in ['datetime64[ns]', '<M8[ns]']:
                    worksheet.write_datetime(row_num + 1, col_num + index_col_offset, value,
                                             self.cell_format[('date', depth, format_option)])
                else:
                    worksheet.write_string(row_num + 1, col_num + index_col_offset, str(value),
                                           self.cell_format[('default', depth, format_option)])

            # Set optional grouping of rows
            if group_rows:
                if depth > 0:
                    worksheet.set_row(row_num + 1, None, None, {'level': depth})

        # Autofit column width
        for col_num, width in enumerate(self.__get_col_widths(output_df)):

            # After the index column, check type and override width if necessary
            if col_num > 0:
                if output_df.dtypes[col_num-1] in ['float64']:
                    width = 8
                elif output_df.dtypes[col_num-1] in ['datetime64[ns]']:
                    width = 8

            # If not printing index, skip to the first column and offset
            if not print_index:
                if col_num == 0: continue
                col_num -= 1

            worksheet.set_column(col_num, col_num, width + 2)


    def write_book(self):
        """ Writes workbook to file after all sheets are added. """
        #self.writer.save()
        self.workbook.close()


    def add_raw_sheet(self, df, sheet_name):
        """ Add a sheet with default pandas formatting """
        #df.to_excel(self.writer, sheet_name=sheet_name, header=True, index=True)
        pass


    def __get_col_widths(self, df):
        """ Return max lengths for each column in DF """

        # First we find the maximum length of the index column
        idx_max = max([len(str(s)) for s in df.index.values])  # + [len(str(dataframe.index.name))])
        # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
        # return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]
        return [idx_max] + [max([len(str(s)) for s in df[col].values]) for col in df.columns]