import pandas as pd
import numpy as np
import easygui
import os
import datetime

import warnings
warnings.filterwarnings("ignore", 'This pattern has match groups')

class BOM:

    def __init__(self, file_path = None):
        if file_path:
            self.file_path = file_path
            self.load_csv(file_path)


    def load_csv(self, file_path = None):
        """ Load CSV from given path, or if None given, prompt user using GUI."""

        if file_path is None:
            self.file_path = easygui.fileopenbox(msg='Choose requirement spreadsheet', default='*.csv',
                                                 filetypes=[["*.csv", "All files"]])
        else:
            self.file_path = file_path

        self.__read_csv_file()
        self.raw_df = self.df.copy()

        # Check for level column:
        self.__check_cols(['Level'])

        # Clean up part numbers:
        self.__process_part_numbers()
        self.__determine_part_type()
        self.__more_stuff()

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
        self.df['Part Number'].replace(to_replace=r"^([1,2][0-9]{2}[F,Q,N,G,E,X][0-9]{4}[-][0-9]*).*",
                                       value=r"\1", regex=True, inplace=True)


    def __determine_part_type(self):
        """Determine type of item (DSS part/assy or COTS) using some regex magic"""

        dss_part_filter = self.df['Part Number'].str.contains('[1,2][0-9]{2}[F,Q,N,G,E,X][0-9]{4}', na=False)
        self.df.loc[dss_part_filter & (self.df['Extension'] == 'SLDPRT'), 'Type'] = 'DSS PART'
        self.df.loc[dss_part_filter & (self.df['Extension'] == 'SLDASM'), 'Type'] = 'DSS ASSY'
        self.df.loc[~dss_part_filter, 'Type'] = 'COTS'


    def __more_stuff(self):
        # Determine drawing number from valid DSS items
        dss_drw_filter = self.df['Part Number'].str.contains('^[1,2][0-9]{2}[F,Q,N,G,E,X][0-9]{4}', na=False)
        self.df['Drawing Number'] = self.df['Part Number'].loc[dss_drw_filter].replace(
            to_replace=r"^([1,2][0-9]{2}[F,Q,N,G,E,X][0-9]{4}).*", value=r"\1", regex=True)

        # Determine if drawing (i.e. DSS Part or Assembly is a -1 number and is not duplicate)
        dash1_filter = self.df['Part Number'].str.contains('[1,2][0-9]{2}[F,Q,N,G,E,X][0-9]{4}(-1$|-1_)', na=False)
        self.df.loc[dash1_filter & (~self.df.duplicated('Part Number', 'first')), 'Drawing'] = 'Yes'

        # Mark duplicate parts
        self.df.loc[self.df.duplicated('Part Number', 'first'), 'Duplicate'] = 'Yes'

        # Determine depth of line (count level dots)
        self.df['Depth'] = self.df.loc[~self.df['Level'].isnull(), 'Level'].astype('str').str.count('[.]')

        # Assign N/A for Material on Assemblies
        self.df.loc[(self.df['Material'].isnull()) & (self.df['Extension'] == 'SLDASM'), 'Material'] = 'N/A - Assembly'

        # The following code determines "Total QTY" of a line item based on parent assembly quantities
        # "Used On" is also determined

        self.df['Total QTY'] = self.df['QTY']

        for row_num, row in self.df.iterrows():
            row_depth = row['Depth']

            # Filter all rows after current with the same depth
            same_depth_rows = self.df.loc[row_num + 1:].loc[self.df['Depth'] <= row_depth]

            # Get index of next row at same depth or set to None
            next_row = same_depth_rows.head().index[0] if len(same_depth_rows) else 0

            self.df.loc[row_num, 'Next Row'] = next_row

            # Multiply all lower rows by current row QTY
            if next_row > row_num + 1:
                self.df.loc[row_num + 1: next_row - 1, 'Total QTY'] *= row['QTY']
                self.df.loc[row_num + 1: next_row - 1, 'Used On'] = row['Part Number']


    def __check_cols(self, required_cols = []):
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
