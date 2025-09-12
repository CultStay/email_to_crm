from odoo import models, fields, api

class StockPicking(models.Model):
    _inherit = 'stock.picking'

    source_location_id = fields.Many2one('stock.location', string='Source Location', compute='_compute_source_location', store=True)

    def _compute_source_location(self):
        for picking in self:
            sales_man = self.env['stock.location'].search([('name', '=', 'SALES MAN')], limit=1)
            if sales_man:
                picking.source_location_id = sales_man.id
                picking.location_id = sales_man.id
            else:
                picking.source_location_id = False