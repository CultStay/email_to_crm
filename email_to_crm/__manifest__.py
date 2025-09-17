{
    'name': 'Email CRM Lead generation',
    'version': '18.0.0.0.1',
    'category': 'Connector',
    'author': 'Sudarsanan P.R',
    'website': '',
    'summary': 'Connect Email and create leads when a new Booking has come',
    'description': """Email CRM integration""",
    'depends': ['base', 'contacts', 'crm', 'stock', 'sale', 'mail', 'account','sale_crm', 'purchase', 'sale_management'],
    'data': [
        'security/ir.model.access.csv',
        'security/security.xml',
        'views/report_invoice_document.xml',
        'views/fetch_mail.xml',
        'views/crm_lead_inherit.xml',
        'wizard/create_invoice_wizard.xml',
        'views/menu.xml',
    ],

    'license': 'OPL-1',
    'application': False,
    'auto_install': False,
    'installable': True,
}
