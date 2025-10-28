from odoo import models, fields, api

class StockPicking(models.Model):
    _inherit = 'stock.picking'

    @api.model_create_multi
    def create(self, vals):
        sales_man = self.env['stock.location'].search([('name', '=', 'SALES MAN')], limit=1)
        for val in vals:
            if sales_man and not val.get('location_id'):
                val['location_id'] = sales_man.id
        return super().create(vals)
