ODOO_URL = 'dss-space.odoo.com'
ODOO_DB = 'dss-space'
ODOO_USERNAME = 'pdm@dss-space.com'
ODOO_UID = 17
ODOO_PASSWORD = 'odootothemoon!!!'

import odoorpc
import easygui
from datetime import datetime as dt
import pandas as pd
import dfexporter
from ast import literal_eval
import dfexporter

class OdooLoader():

    def __init__(self, srv=ODOO_URL, db=ODOO_DB, user=ODOO_USERNAME, pwd=ODOO_PASSWORD):
        self.api = odoorpc.ODOO(srv, protocol='jsonrpc+ssl', port=443)
        self.api.login(db, user, pwd)
        self.uid = self.api.env.uid


    def search_by_field(self, model, search_field=None, search_string=None):
        Model = self.api.env[model]
        domain = [(search_field,'ilike',search_string)] if search_field else []
        fields = None
        return Model.search_read(domain, fields)

    def get_raw_tasks_df(self):
        return pd.DataFrame(self.search_by_field('project.task'))

    def get_po_headers_df(self):
        return pd.DataFrame(self.search_by_field('purchase.order'))

    def get_raw_po_lines_df(self):
        return pd.DataFrame(self.search_by_field('purchase.order.line'))

    def get_po_lines_df(self, all_jobs=False):

        df = self.get_raw_po_lines_df()

        # Rename ugly x_studio field for Tasks to task_id
        df.rename(columns={'x_studio_field_zGWBJ': 'task_id'}, inplace=True)

        # Replace tasks showing as False to empty list
        df.loc[df['task_id'] == False, 'task_id'] = \
            df.loc[df['task_id'] == False, 'task_id'].replace(False,'[0, \'Empty\']').apply(lambda x: literal_eval(x))

        # Split list data that matters
        df[['partner_id', 'partner_name']] = pd.DataFrame(df.partner_id.values.tolist(), index=df.index)
        df[['product_id', 'product_name']] = pd.DataFrame(df.product_id.values.tolist(), index=df.index)
        df[['task_id', 'task_name']] = pd.DataFrame(df.task_id.values.tolist(), index=df.index)
        df[['order_id', 'order_name']] = pd.DataFrame(df.order_id.values.tolist(), index=df.index)

        # Extract just the PO number
        df['order_name'] = df['order_name'].str.extract(r'(PO-[0-9]+)')

        # Splits product_name (i.e. [####] DESC) into two components:
        df['product_number'] = df.product_name.str.extract(r'\[(.*)\] (.*)', expand=True).loc[:, 0]
        df['product_description'] = df.product_name.str.extract(r'\[(.*)\] (.*)', expand=True).loc[:, 1]

        column_map = {
            'order_name':'PO Number',
            'partner_name':'Supplier',
            'x_studio_line_':'PO Line Number',
            'state':'Status',
            'task_name':'Job ID',
            'product_number':'Product Number',
            'x_studio_po_revision':'Product Revision',
            'product_description':'Description',
            'product_uom_qty':'Qty Ordered',
            'qty_received':'Qty Recd',
            'date_planned':'Due Date',
            'price_unit':'Unit Price',
            'price_tax':'Tax Price',
            'price_total':'Total Price'
        }
        df.rename(columns=column_map, inplace=True)

        column_list = [
            'PO Number',
            'Supplier',
            'PO Line Number',
            'Status',
            'Job ID',
            'Product Number',
            'Product Revision',
            'Description',
            'Qty Ordered',
            'Qty Recd',
            'Due Date',
            'Unit Price',
            'Tax Price',
            'Total Price'
        ]

        # Prompt and filter PO DF for Job list
        if all_jobs is False: df = self.df_job_filter(df)

        return df[column_list]

    def get_lots_df(self):
        return pd.DataFrame(self.search_by_field('stock.production.lot'))

    def get_stock_report_df(self):
        return pd.DataFrame(self.search_by_field('stock.report'))


    def export_all_data(self):
        export_file_name = f'{int(dt.now().timestamp())} odoo data.xlsx'
        excel_export = dfexporter.DFExport(export_file_name)

        excel_export.add_sheet(self.get_lots_df(),'Lots')
        excel_export.add_sheet(self.get_po_headers_df(), 'PO Headers')
        excel_export.add_sheet(self.get_po_lines_df(), 'PO Lines')
        excel_export.add_sheet(self.get_stock_report_df(), 'Stock Report')
        excel_export.write_book()

    def df_job_filter(self, df, jobs=None):
        """ Filters PO data DF by 'Job ID' with given list, or if none, prompts user. Returns filtered DF. """
        task_list = self.get_raw_tasks_df().sort_values('sequence').name
        if jobs is None:
            jobs = easygui.multchoicebox('Which Odoo tasks/jobs do you want included?',
                                         choices=task_list.unique())
        return df.loc[df['Job ID'].isin(jobs)]


def test():
    global odoo, df
    odoo=OdooLoader()

    # Export all DF to excel
    # odoo.export_all_data()

    # Get DF of just PO line data
    df = odoo.get_po_lines_df()

def export(df):
    excel_export = dfexporter.DFExport('odoo_po_data.xlsx')
    odoo_po_cols = [
        'PO Number',
        'Supplier',
        'PO Line Number',
        'Status',
        'Job ID',
        'Product Number',
        'Product Revision',
        'Description',
        'Qty Ordered',
        'Qty Recd',
        'Due Date',
        'Unit Price',
        'Tax Price',
        'Total Price'
    ]
    excel_export.add_sheet(df,
                                        sheet_name='Odoo PO Data',
                                        cols_to_print=odoo_po_cols,
                                        print_index=False)
    excel_export.write_book()


if __name__ == 'builtins':
    test()
