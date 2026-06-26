from odoo import models, fields, api


class SaleOrder(models.Model):
    _inherit = 'sale.order'

    city = fields.Char(
        string="City",
        related="partner_id.city",
        store=True
    )

    payment_status = fields.Selection(
        [
            ('unpaid', 'Unpaid'),
            ('partial', 'Partial Payment'),
            ('paid', 'Paid'),
        ],
        string="Payment Status",
        compute="_compute_payment_status",
        store=True,
    )

    @api.depends('invoice_ids.payment_state', 'invoice_ids.state')
    def _compute_payment_status(self):
        for order in self:
            invoices = order.invoice_ids.filtered(lambda inv: inv.state == 'posted')

            if not invoices:
                order.payment_status = 'unpaid'
                continue

            states = invoices.mapped('payment_state')

            if all(state == 'paid' for state in states):
                order.payment_status = 'paid'
            elif any(state in ('partial', 'in_payment') for state in states):
                order.payment_status = 'partial'
            else:
                order.payment_status = 'unpaid'
