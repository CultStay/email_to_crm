# -*- coding: utf-8 -*-
{
    'name': 'Custom Report',
    'version': '18.0.1.0.0',
    'category': 'Sales/Reports',
    'summary': 'Sales / Invoice Report',
    'description': """
        Report Module
        ========================
        Generates Excel report of sales orders/invoices.
        - Filter by date range via wizard
        - Export to Excel (.xlsx)
    """,
    'author': 'Custom',
    'depends': ['sale_management', 'account'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/collection_report_wizard_view.xml',
        'wizard/return_report_wizard_view.xml',
    ],
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
}
