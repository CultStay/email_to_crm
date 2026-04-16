from odoo import models, fields, api


class SaleOrder(models.Model):
    _inherit = 'sale.order'

    city = fields.Char(
        string="City",
        related='partner_id.city',
        store=True
    )