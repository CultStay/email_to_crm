{
    'name': 'Email CRM Lead generation',
    'version': '18.0.0.0.1',
    'category': 'Connector',
    'author': 'Sudarsanan P.R',
    'website': '',
    'summary': 'Connect Email and create leads when a new Booking has come',
    'description': """Email CRM integration""",
    # NOTE: 'l10n_in' pins this module to databases with the Indian localization
    # installed (consistent with the INR amounts / TDS fields used in the mail
    # parsing logic). No functional issue found here — left unchanged from the
    # original. If this module needs to run on a non-Indian company, this
    # dependency will block installation.
    'depends': ['base','web', 'contacts', 'crm', 'stock', 'sale', 'mail', 'account','sale_crm', 'purchase', 'sale_management', 'l10n_in'],
    'data': [
        'security/ir.model.access.csv',
        'security/security.xml',
        'report/daily_statement_report.xml',
        'views/report_invoice_document.xml',
        'views/fetch_mail.xml',
        'views/account_move.xml',
        'views/crm_lead_inherit.xml',
        'wizard/create_invoice_wizard.xml',
        'views/menu.xml',
        'views/stock_location.xml',
        'data/ir_cron.xml',
        'views/res_config_settings.xml',  
    ],
    # NOTE: worth double-checking that 'security/ir.model.access.csv' grants the
    # user/group under which incoming mail is fetched (e.g. the fetchmail cron
    # user) create rights on res.partner, account.move, and account.payment —
    # message_process() creates records on these models directly, and a missing
    # access rule here would raise a silent AccessError for whichever provider's
    # flow touches the restricted model first, which can look identical to a
    # parsing bug from the outside.

    'license': 'OPL-1',
    'application': False,
    'auto_install': False,
    'installable': True,
}
