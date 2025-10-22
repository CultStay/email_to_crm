from odoo import fields, models

class ResConfigSettings(models.TransientModel):
    _inherit = 'res.config.settings'

    crm_report_email = fields.Char(string="CRM Report Email", config_parameter='crm.report_email')
