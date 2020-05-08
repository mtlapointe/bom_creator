""" Module used to load BOM CSV file from PDM into a Pandas DataFrame

    Typical usage example:
    TBD
"""

import pandas as pd
import numpy as np
import easygui
import os
import datetime
import math
import warnings

warnings.filterwarnings("ignore", 'This pattern has match groups')


class BOM:
    """ BOM class loads CSV BOM file and converts into a clean Pandas DataFrame

        Attributes:
            file_path: File path of CSV file
            raw_df: DataFrame with unprocessed CSV data
    """

    def __init__(self, file_path=None):
        """ Constructor for class, can optionally load CSV from given file path.

        Args:
            file_path (str, optional): File path for CSV file to load
        """
        if file_path:
            self.file_path = file_path
            self.load_csv(file_path)

    def load_csv(self, file_path=None):
        """ Load CSV from given path, or if None given, prompt user using GUI.

        Args:
            file_path (str, optional): File path for CSV file to load

        Returns:
            Self instance of class object
        """

        # Prompt user for file path if none given
        if file_path is None:
            self.file_path = easygui.fileopenbox(msg='Choose PDM BOM CSV file', default='*.csv',
                                                 filetypes=[["*.csv", "All files"]])
        else:
            self.file_path = file_path

        # Read CSV file into DataFrame. Throw error if trying to read non-CSV
        if '.csv' not in self.file_path:
            raise RuntimeError("Trying to load a non-CSV file...")

        self.df = pd.read_csv(self.file_path, encoding='utf_16', dtype={'Level': object},
                              float_precision='round_trip', error_bad_lines=False)

        # Create copy of DataFrame for raw data
        self.raw_df = self.df.copy()

        # Create a unique Unique ID for each line
        self.df = self.df.reset_index().rename(columns={'index': 'Unique ID'})

        # Check for level column:
        self.__check_cols(['Level'])

        # Clean up part numbers:
        self.__process_part_numbers()
        self.__determine_part_type()
        self.__sort_df()
        self.__get_used_on()
        self.__get_total_qty()
        self.__more_stuff()

        # Cast specific data types
        Int64 = pd.Int64Dtype()

        data_types = {'Depth': Int64,
                      'ID': Int64,
                      'Latest Version': Int64,
                      'QTY': Int64,
                      'Unique ID': Int64,
                      'Parent ID': Int64,
                      'Total QTY': Int64}

        self.df = self.df.astype(data_types)

        return self

    def __read_csv_file(self):
        """ Internal method to read CSV file and load into DF"""
        if '.csv' not in self.file_path:
            raise RuntimeError("Trying to load a non-CSV file...")

        self.df = pd.read_csv(self.file_path, encoding='utf_16', dtype={'Level': object},
                              float_precision='round_trip', error_bad_lines=False)

    def __process_part_numbers(self):
        """ Determine part number from file name and config, or part number override"""

        self.__check_cols(['Name', 'Configuration', 'PartNumOverride'])

        # Fix filename - force uppercase and split off extension
        self.df['File Name'], self.df['Extension'] = self.df['Name'].str.strip().str.upper().str. \
            rsplit('.', n=1).str
        # Drop any non-SW file from the list (gets rid of PSELF.DFs, etc)
        self.df = self.df[self.df['Extension'].isin(['SLDPRT', 'SLDASM'])].reset_index(drop=True)

        # Remove data from Part Number Column (crap data from PDM...)
        self.df['Part Number'] = np.NaN

        # Check for NOCONFIG and assign File Name only to Part Number
        self.df.loc[self.df['Configuration'].fillna('NOCONFIG').str.upper() == 'NOCONFIG',
                    ['Part Number']] = self.df['File Name']

        # Check for PN Override and use that if true
        self.df.loc[self.df['PartNumOverride'].notnull(), ['Part Number']] = self.df['PartNumOverride']

        # Everything else, set PN to FILENAME + CONFIG
        self.df.loc[self.df['Part Number'].isnull(), ['Part Number']] = self.df['File Name'] + self.df['Configuration']

        # For DSS PNs, strip out anything after the dash number (ex. 100-DEPLOYED)
        # https://stackoverflow.com/a/41609175/6475884 <- how the regex replace works
        self.df['Part Number'].replace(to_replace=r"^([1,2][0-9]{2}[F,Q,N,G,E,X,T][0-9]{4}[-][0-9]*).*",
                                       value=r"\1", regex=True, inplace=True)

    def __determine_part_type(self):
        """Determine type of item (DSS part/assy or COTS) using some regex magic"""

        dss_part_filter = self.df['Part Number'].str.contains('[1,2][0-9]{2}[F,Q,N,G,E,X,T][0-9]{4}', na=False)
        self.df.loc[dss_part_filter & (self.df['Extension'] == 'SLDPRT'), 'Type'] = 'DSS PART'
        self.df.loc[dss_part_filter & (self.df['Extension'] == 'SLDASM'), 'Type'] = 'DSS ASSY'
        self.df.loc[~dss_part_filter, 'Type'] = 'COTS'

    def __more_stuff(self):
        # Determine drawing number from valid DSS items
        dss_drw_filter = self.df['Part Number'].str.contains('^[1,2][0-9]{2}[F,Q,N,G,E,X,T][0-9]{4}', na=False)
        self.df['Drawing Number'] = self.df['Part Number'].loc[dss_drw_filter].replace(
            to_replace=r"^([1,2][0-9]{2}[F,Q,N,G,E,X,T][0-9]{4}).*", value=r"\1", regex=True)

        # Determine if drawing (i.e. DSS Part or Assembly is a -1 number and is not duplicate)
        dash1_filter = self.df['Part Number'].str.contains('[1,2][0-9]{2}[F,Q,N,G,E,X,T][0-9]{4}(-1$|-1_)', na=False)
        self.df.loc[dash1_filter & (~self.df.duplicated('Part Number', 'first')), 'Drawing'] = 'Yes'

        # Mark duplicate parts
        self.df.loc[self.df.duplicated('Part Number', 'first'), 'Duplicate'] = 'Yes'

        # Assign N/A for Material on Assemblies
        self.df.loc[(self.df['Material'].isnull()) & (self.df['Extension'] == 'SLDASM'), 'Material'] = 'N/A - Assembly'

    def __sort_df(self):

        df = self.df

        # First, figure out depth (i.e. count how many dots are in level)
        df['Depth'] = df.loc[df['Level'].notnull(), 'Level'].astype('str').str.count('[.]')

        # Second, determine the parent item Level (i.e. drop off last number, ex. 1.3.2.1 becomes 1.3.2
        parent_level = df.loc[df['Level'].notnull(), 'Level'].astype('str').str.split('.').apply(
            lambda x: '.'.join(x[:-1]))

        # Look-up Unique ID for given Level
        def lookup_parent_id(x):
            return int(df.loc[df['Level'] == x]['Unique ID'].iloc[0]) if (x and x is not np.nan) else None

        df['Parent ID'] = parent_level.apply(lookup_parent_id)

        # Third, sort values at top level (i.e. depth == 0) and create initial DF
        sorted_df = df.loc[df['Depth'] == 0].sort_values(by='Part Number').reset_index(drop=True)
        sorted_df['New Level'] = sorted_df.index + 1

        # Group the rest of the DF by the depth and then the parent values, sort the groups by the name column
        for key, group_df in df.loc[df['Depth'] > 0].sort_values(by='Part Number').groupby(by=['Depth', 'Parent ID']):
            # key = [Depth, Parent ID]
            # group_df = DataFrame for that group

            # Reset the index order of the group
            sorted_group_df = group_df.reset_index(drop=True)

            # Get the new parent level
            parent_new_level = str(sorted_df.loc[sorted_df['Unique ID'] == key[1], 'New Level'].iloc[0])

            # group_df['New Level'] = group_df.index.to_series().astype(int).apply(
            #    lambda x: '.'.join([parent_new_level, str(x + 1)]))
            sorted_group_df['New Level'] = sorted_group_df.index.to_series().apply(
                lambda x: '.'.join([parent_new_level, str(int(x) + 1)]))

            # In the new DataFrame, get the index location of the Parent for this group, then split the DF into two
            parent_index = sorted_df.loc[sorted_df['Unique ID'] == key[1]].index[0] + 1
            start_slice = sorted_df.iloc[0:parent_index]
            end_slice = sorted_df.iloc[parent_index:]

            # Insert the new sorted group in the correct location in the new DF (after the parent)
            sorted_df = pd.concat([start_slice, sorted_group_df, end_slice], sort=True).reset_index(drop=True)

        self.df = sorted_df.rename(columns={'New Level': 'Level', 'Level': 'Old Level'})

    def __get_used_on(self):

        def lookup_parent(x):
            return self.df.loc[self.df['Unique ID'] == x]['Part Number'].iloc[0] if not math.isnan(x) else None

        self.df['Used On'] = self.df['Parent ID'].apply(lookup_parent)

        def get_parent_list(x):
            parent_list = []
            next_parent = x
            while math.isnan(next_parent) is False:
                parent_list.append(int(next_parent))
                next_parent = self.df.loc[self.df['Unique ID'] == next_parent]['Parent ID'].iloc[0]
            return parent_list

        self.df['Parent List'] = self.df['Parent ID'].apply(get_parent_list)

    def __get_total_qty(self):

        top_level_qty = self.df['QTY'][0]  # Probably always 1?

        def get_parent_qtys(parent_id_list):
            total_qty = top_level_qty
            for id in parent_id_list:
                total_qty *= self.df.loc[self.df['Unique ID'] == id]['QTY'].iloc[0]
            return total_qty

        parent_qtys = self.df['Parent List'].apply(get_parent_qtys)

        self.df['Total QTY'] = self.df['QTY'] * parent_qtys

    def __check_cols(self, required_cols=[]):
        """ Checks if list of required columns is in CSV header"""

        missing_cols = np.setdiff1d(required_cols, self.df.columns)
        if missing_cols.size > 0:
            raise RuntimeError(f'CSV is missing the following columns: {missing_cols}')

    def get_assy_from_file(self):
        # Strip out extension, ex. 1234567.SLDASM.1.BOM
        csv_file_name = os.path.splitext(os.path.basename(self.file_path))[0]
        # Extract assembly number
        return csv_file_name.split('.')[0]

    def get_date_from_file(self):
        return datetime.datetime.fromtimestamp(os.path.getatime(self.file_path)).strftime('%Y%m%d')


def main():
    global bom
    bom = BOM().load_csv()
    export()


def export():
    import dfexporter
    from datetime import datetime as dt

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

    excel_export = dfexporter.DFExport(f'{int(dt.now().timestamp())} BOM Test.xlsx')

    excel_export.add_sheet(bom.df,
                           sheet_name='Assembly BOM',
                           freeze_col=5, freeze_row=1,
                           cols_to_print=full_bom_cols,
                           depth_col_name='Depth',
                           group_rows=True,
                           highlight_depth=True,
                           highlight_col_limit=0,
                           cols_to_indent=['Part Number'],
                           print_index=True)

    excel_export.write_book()


def test():
    global bom

    bom = BOM(file_path='test/test.csv')


if __name__ == 'builtins':
    test()
