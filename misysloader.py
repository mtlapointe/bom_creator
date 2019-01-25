import pandas
import pyodbc
import dfexporter

class MisysTable:

    def __init__(self):
        self.server = '192.168.75.21,1500'
        self.database = 'DSS'


    def load_sql(self, sql):
        """ Connect to MISys DB, run SQL query and return results as DF """

        # Connect to DB
        cnxn = pyodbc.connect(r'DRIVER={ODBC Driver 17 for SQL Server};'
                              f'SERVER={self.server};'
                              f'DATABASE={self.database};'
                              r'TRUSTED_CONNECTION=YES')
        #cnxn.setencoding('utf-8')

        # Run SELECT sql query and load into DF
        df = pandas.read_sql(sql, cnxn)

        # rowVer is an oddball column that causes encoding errors - BE GONE!
        if 'rowVer' in df.columns:
            df.drop(columns=['rowVer'], inplace=True)
        return df

    def load_po_data(self, row_limit=None):
        """ Canned SQL that gets PO line item data """
        sql = (f'SELECT {f"TOP {row_limit}" if row_limit else ""} '
               'MIPOD.[pohId] AS [PO Number], '
               'MIPOH.[name] AS [Supplier], '
               'MIPOD.[lineNbr] AS [PO Line Numer], '
               'MIPOD.[dType] AS [Data Type], '
               'MIPOD.[dStatus] AS Status, '
               'MIPOD.[jobId] AS [Job ID], '
               'MIPOD.[locId] AS [Location ID], '
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
               'MIPOD.[lastRecvDt] AS [Date Last Recd] '
               'FROM MIPOH RIGHT JOIN MIPOD ON MIPOH.pohId = MIPOD.pohId '
               'ORDER BY MIPOD.[pohId] DESC, MIPOD.[lineNbr] ASC')
        return self.load_sql(sql)

    def load_table(self, table, row_limit=None):
        """ Runs SELECT * from specified table. Can limit number of rows. """
        df = self.load_sql(f'SELECT {f"TOP {row_limit}" if row_limit else ""} * FROM {table}')
        return df

    def load_raw_po_df(self, row_limit=None):
        """ Loads PO Line and PO Header tables in DF's, then does a pandas join. Slow than a SQL join, but don't
        have to deal with conflicting header column names """
        mipoh = misys.load_table('MIPOH')
        mipod = misys.load_table('MIPOD',row_limit)
        df = mipoh.join(mipod.set_index('pohId'), on='pohId', lsuffix='_mipod', rsuffix='_mipoh', how='right') \
            .reset_index()
        return df

    def export_raw_po_df(self, file_name='raw_po_data.xlsx', row_limit=None):
        """ Dump raw PO line item data to XLSX with given name. Can also limit number of rows """
        self.load_raw_po_df(row_limit).to_excel(file_name)



def example():
    """ Example use """
    misys = MisysTable()
    df = misys.load_po_data(2500)
    export(df)


def export(dataframe):
    """ Exports DF to Excel using pretty format """
    excel = dfexporter.DFExport('misys.xlsx')
    excel.add_sheet(dataframe)
    excel.write_book()