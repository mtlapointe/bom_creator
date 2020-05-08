""" NOTE - You need to download the Microsoft ODBC Driver 17 for this code to work
https://www.microsoft.com/en-us/download/details.aspx?id=56567
https://www.microsoft.com/en-us/sql-server/developer-get-started/python/windows/"""

import pandas as pd
import pyodbc
import dfexporter
import numpy as np
import os
import warnings
import datetime
import easygui


class MisysTable:

    def __init__(self, force_update=False, cache_age_limit=24):
        self.server = '192.168.75.21,1500'
        self.database = 'DSS'
        self.cache_dir = 'cache'
        self.cache_age_limit = cache_age_limit
        self.username = 'exporter'
        self.password = 'password'
        self.force_update = force_update

    def load_sql(self, sql, cache_name):
        """ Connect to MISys DB, run SQL query and return results as DF """

        if self.force_update or not self.check_for_cache(cache_name) \
                or self.cache_age(cache_name) > self.cache_age_limit:
            try:
                print('Fetching MISys data from database')
                # Connect to DB
                cnxn = pyodbc.connect(r'DRIVER={ODBC Driver 17 for SQL Server};'
                                      f'SERVER={self.server};'
                                      f'DATABASE={self.database};'
                                      f'UID={self.username};'
                                      f'PWD={self.password};', timeout=5)

                # Run SELECT sql query and load into DF
                df = pd.read_sql(sql, cnxn)

                # rowVer is an oddball column that causes encoding errors - BE GONE!
                if 'rowVer' in df.columns:
                    df.drop(columns=['rowVer'], inplace=True)

                # Replace empty strings with nan
                df.replace('', np.nan, regex=True, inplace=True)

                self.save_cache(df, cache_name)
                return df

            except:
                if self.check_for_cache(cache_name):
                    warnings.warn('Could not connect to DB, using outdated cache data that is '
                                  f'{self.cache_age(cache_name):.2f} hours old.')
                    return self.read_cache(cache_name)
                else:
                    raise Exception('Cannot read data from DB or cache!!')

        elif self.check_for_cache(cache_name):
            print('Fetching MISys data from cache')
            return self.read_cache(cache_name)

        else:
            raise Exception('Cannot read data from DB or cache!!')

    def fetch_po_data(self, row_limit=None):
        """ Canned SQL that gets PO line item data from MIPOH and MIPOD tables """
        sql = (f'SELECT {f"TOP {row_limit}" if row_limit else ""} '
               'MIPOD.[pohId] AS [PO Number], '
               'MIPOH.[name] AS [Supplier], '
               'MIPOD.[lineNbr] AS [PO Line Number], '
               'MIPOD.[dStatus] AS Status, '
               'MIPOD.[jobId] AS [Job ID], '
               'MIPOD.[itemId] AS [Item Number], '
               'MIPOD.[viCode] AS [Misc Item Number], '
               'MIPOD.[descr] AS [Description], '
               'MIPOD.[cmt] AS [Comment], '
               'MIPOD.[ordered] AS [Qty Ordered], '
               'MIPOD.[received] AS [Qty Recd], '
               'MIPOD.[poUOfM] AS UOM, '
               'MIPOD.[poXStk] AS [UOM Conversion], '
               'MIPOD.[price] AS [Unit Price], '
               'MIPOD.[initDueDt] AS [Initial Due Date], '
               'MIPOD.[realDueDt] AS [Actual Due Date], '
               'MIPOD.[promisedDt] AS [Promised Date], '
               'MIPOD.[lastRecvDt] AS [Date Last Recd], '
               'MIPOD.[dType] AS [Data Type], '
               'MIPOD.[locId] AS [Location ID] '
               'FROM MIPOH RIGHT JOIN MIPOD ON MIPOH.pohId = MIPOD.pohId '
               'ORDER BY MIPOD.[pohId] DESC, MIPOD.[lineNbr] ASC')

        return self.load_sql(sql, 'PO_TABLE')

    def po_data_job_filter(self, df, jobs=None):
        """ Filters PO data DF by 'Job ID' with given list, or if none, prompts user. Returns filtered DF. """
        if jobs is None:
            jobs = easygui.multchoicebox('Which MISys jobs do you want included?',
                                         choices=df['Job ID'].sort_values().unique())
        return df.loc[df['Job ID'].isin(jobs)]

    def load_po_data(self, filter_jobs=None):
        df = self.fetch_po_data()
        # Join Item Number and Misc Item Number
        # df['Product Number'] = df['Item Number'].combine_first(df['Misc Item Number'])
        df.insert(7, 'Product Number', df['Item Number'].combine_first(df['Misc Item Number']))

        # Regex to split off DSS number from REV or other info
        split_dss_num_regex = '(^[1,2][0-9]{2}[F,Q,N,G,E,X,T][0-9]{4}(?:[-]\w*)*)(?:\s*)(.*)'
        split_dss_number = df['Product Number'].str.extract(split_dss_num_regex)
        df['Product Number'] = split_dss_number[0].fillna(df['Product Number'])
        df.insert(8, 'Product Revision', split_dss_number[1])
        df.drop(['Item Number', 'Misc Item Number'], axis=1, inplace=True)

        df['Qty Ordered'] = df['Qty Ordered'].astype('int64')
        df['Qty Recd'] = df['Qty Recd'].astype('int64')

        # Filter by jobs
        df = self.po_data_job_filter(df)

        return df

    def load_table(self, table, row_limit=None):
        """ Runs SELECT * from specified table. Can limit number of rows. """
        df = self.load_sql(f'SELECT {f"TOP {row_limit}" if row_limit else ""} * FROM {table}', table)
        return df

    def load_raw_po_df(self, row_limit=None):
        """ Loads PO Line and PO Header tables in DF's, then does a pandas join. Slow than a SQL join, but don't
        have to deal with conflicting header column names """
        mipoh = misys.load_table('MIPOH')
        mipod = misys.load_table('MIPOD', row_limit)
        df = mipoh.join(mipod.set_index('pohId'), on='pohId', lsuffix='_mipod', rsuffix='_mipoh', how='right') \
            .reset_index()
        return df

    def export_raw_po_df(self, file_name='raw_po_data.xlsx', row_limit=None):
        """ Dump raw PO line item data to XLSX with given name. Can also limit number of rows """
        self.load_raw_po_df(row_limit).to_excel(file_name)

    def fix_pn(self, df):
        df.replace('', np.nan, regex=True, inplace=True)

    def save_cache(self, df, cache_name):
        cache_dir = f'./{self.cache_dir}'
        if not os.path.exists(cache_dir):
            os.mkdir(cache_dir)
        cache_path = os.path.join(cache_dir, cache_name)
        df.to_pickle(cache_path)

    def read_cache(self, cache_name):
        cache_path = f'{self.cache_dir}/{cache_name}'
        if os.path.exists(cache_path):
            return pd.read_pickle(cache_path)
        else:
            return None

    def check_for_cache(self, cache_name):
        cache_path = f'{self.cache_dir}/{cache_name}'
        return os.path.exists(cache_path)

    def cache_age(self, cache_name):
        cache_path = f'{self.cache_dir}/{cache_name}'
        if os.path.exists(cache_path):
            cache_age = datetime.datetime.now() - datetime.datetime.fromtimestamp(os.path.getmtime(cache_path))
            return cache_age.total_seconds() / (60 * 60)
        else:
            return 0


def example():
    """ Example use """
    misys = MisysTable()
    df = misys.load_po_data()
    export(df)


def export(dataframe):
    """ Exports DF to Excel using pretty format """
    excel = dfexporter.DFExport('misys.xlsx')
    excel.add_sheet(dataframe)
    excel.add_raw_sheet(dataframe, 'raw')
    excel.write_book()


def test():
    for row_num, (_, row) in enumerate(output_df.iterrows()):
        for col_num in range(len(row)):
            misys = MisysTable()
    df = misys.load_po_data()
