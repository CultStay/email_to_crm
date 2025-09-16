from odoo import models, fields, api

class StockPicking(models.Model):
    _inherit = 'stock.picking'

    @api.model
    def create(self, vals):
        sales_man = self.env['stock.location'].search([('name', '=', 'SALES MAN')], limit=1)
        if sales_man and not vals.get('location_id'):
            vals['location_id'] = sales_man.id
        return super().create(vals)