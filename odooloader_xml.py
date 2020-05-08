ODOO_URL = 'https://dss-space-0324-demo.odoo.com'
ODOO_DB = 'dss-space-0324-demo'
ODOO_USERNAME = 'pdm@dss-space.com'
ODOO_UID = 17
ODOO_PASSWORD = 'odootothemoon!!!'

import xmlrpc.client

class OdooLoader:

    def __init__(self):
        pass

    def connect(self):
        try:
            self.common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(ODOO_URL))
            # uid = common.authenticate(ODOO_DB, ODOO_USERNAME, ODOO_PASSWORD, {})
            self.api = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(ODOO_URL))
            return self

        except Exception as e:
            print(e)
            return False


def test():
    global odoo
    odoo = OdooLoader().connect()
    odoo.api.execute_kw(ODOO_DB, ODOO_UID, ODOO_PASSWORD, 'res.partner', 'search_count', [[]])
    tasks = odoo.api.execute_kw(ODOO_DB, ODOO_UID, ODOO_PASSWORD, 'project.task', 'search_read', [[]])
    task_list = [d['name'] for d in tasks]


if __name__ == 'builtins':
    test()
