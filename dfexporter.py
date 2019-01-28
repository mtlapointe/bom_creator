import pandas as pd


class DFExport:

    def __init__(self, output_file_name = "output.xlsx"):
        """ Create BOMExporter object. Use given output file name or default. """

        self.output_file_name = output_file_name

        # Setup Excel writer
        self.writer = pd.ExcelWriter(self.output_file_name, engine='xlsxwriter',
                                     options={'nan_inf_to_errors': True})

        self.workbook = self.writer.book

        self.__load_default_style()


    def __load_default_style(self):
        """ Setup default sheet styles for xlsxwriter """

        default_style = {
            'bold': False,
            'border': 1
        }

        header_style = {
            'rotation': 90,
            'align': 'Center',
            'bold': True
        }

        index_style = {
            'bold': True,
            'align': 'Center',
        }

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

        # Assign formats for default, header and index
        self.default_format = self.workbook.add_format({**default_style})
        self.header_format = self.workbook.add_format({**default_style, **header_style})
        self.index_format = self.workbook.add_format({**default_style, **index_style})

        # Assign formats for highlighted cells and index
        self.highlight_format = []
        self.index_highlight_format = []
        for format in depth_colors:
            self.highlight_format.append(self.workbook.add_format({**default_style, **format}))
            self.index_highlight_format.append(self.workbook.add_format({**default_style, **index_style, **format}))

            # Assign formats for indented cells and indented+highlighted cells
        self.indent_format = []
        self.indent_highlight_format = []
        for i in range(0, 10):
            self.indent_format.append(self.workbook.add_format({**default_style, **{'indent': i}}))
            self.indent_highlight_format.append(self.workbook.add_format({**default_style, **{'indent': i}}))
            if i < len(depth_colors):
                self.indent_highlight_format[i].set_bg_color(depth_colors[i]['bg_color'])
                self.indent_highlight_format[i].set_font_color(depth_colors[i]['font_color'])


    def add_sheet(self, df, sheet_name="Sheet1", zoom=85, freeze_row=1, freeze_col=0, cols_to_print=None,
                  depth_col_name='', cols_to_indent=None, highlight_depth=False, highlight_col_limit=0, group_rows=False,
                  print_index=True):
        """ Take DF and creates new sheet with various options. """

        # Create output DF with only cols to print and replace N/A with empty string
        if cols_to_print:
            output_df = df[cols_to_print].where((pd.notnull(df)), '')
        else:
            output_df = df.where((pd.notnull(df)), '')

        # If index column exists, need offset to shift all other columns
        index_col_offset = 1 if print_index else 0

        # Write data to Excel
        output_df.to_excel(self.writer, sheet_name=sheet_name, startrow=1, startcol=index_col_offset,
                           header=False, index=False)

        # Setup workbook and worksheet objects
        worksheet = self.writer.sheets[sheet_name]

        # Set zoom and freeze panes location
        worksheet.set_zoom(zoom)
        worksheet.freeze_panes(freeze_row, freeze_col)

        # Write the column headers with the defined format.
        if print_index:
            worksheet.write(0, 0, 'Index', self.header_format)
        for col_num, value in enumerate(output_df.columns.values):
            worksheet.write(0, col_num + index_col_offset, value, self.header_format)

        # Iterate through DF rows and write to Excel file
        for row_num, (_, row) in enumerate(output_df.iterrows()):

            # Get the row depth (if needed for highlight, indent or grouping)
            if highlight_depth or cols_to_indent or group_rows:
                depth = int(df[depth_col_name].iloc[row_num])

            # Write optional index first using highlighted or plain index format
            print_format = self.index_highlight_format[depth] if highlight_depth else self.index_format
            if print_index:
                worksheet.write(row_num + 1, 0, output_df.index[row_num], print_format)

            # Write rest of the row
            for col_num in range(len(row)):

                # Check if column should be highlighted and/or indented
                indent_col = cols_to_indent is not None and output_df.columns[col_num] in cols_to_indent
                highlight_col = highlight_depth and \
                                (highlight_col_limit==0 or col_num < highlight_col_limit-index_col_offset)

                # Choose the correct format to use
                if indent_col and highlight_col:
                    print_format = self.indent_highlight_format[depth]
                elif indent_col:
                    print_format = self.indent_format[depth]
                elif highlight_col:
                    print_format = self.highlight_format[depth]
                else:
                    print_format = self.default_format

                # Write data as number or string
                if output_df.dtypes[col_num] == 'int64':
                    worksheet.write_number(row_num + 1, col_num + index_col_offset, row[col_num], print_format)
                else:
                    worksheet.write_string(row_num + 1, col_num + index_col_offset, str(row[col_num]), print_format)

            # Set optional grouping of rows
            if group_rows:
                if depth > 0:
                    worksheet.set_row(row_num + 1, None, None, {'level': depth})

        # Autofit column width
        for i, width in enumerate(self.__get_col_widths(output_df)):
            worksheet.set_column(i+index_col_offset-1, i+index_col_offset-1, width + 2)


    def write_book(self):
        """ Writes workbook to file after all sheets are added. """
        self.writer.save()


    def add_raw_sheet(self, df, sheet_name):
        """ Add a sheet with default pandas formatting """
        df.to_excel(self.writer, sheet_name=sheet_name, header=True, index=True)


    def __get_col_widths(self, df):
        """ Return max lengths for each column in DF """

        # First we find the maximum length of the index column
        idx_max = max([len(str(s)) for s in df.index.values])  # + [len(str(dataframe.index.name))])
        # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
        # return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]
        return [idx_max] + [max([len(str(s)) for s in df[col].values]) for col in df.columns]