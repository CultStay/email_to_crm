from odoo import models, fields, api, _
from odoo.exceptions import UserError

class CreateInvoiceWizard(models.TransientModel):
    _name = 'create.invoice.wizard'
    _description = 'Create Invoice Wizard'

    partner_id = fields.Many2one('res.partner', string='Customer', readonly=True)
    rate = fields.Float(string='Rate', readonly=False)
    customer_paid = fields.Float(string='Customer Paid', required=True)
    lead_id = fields.Many2one('crm.lead', string='Lead')
    property_product_id = fields.Many2one('product.product', string='Product')
    
    def action_create_invoice(self):
        """Create an invoice for the lead."""
        self.ensure_one()
        if not self.partner_id:
            raise UserError(_("Please set a partner before creating an invoice."))
        
        if self.customer_paid <= 0:
            raise UserError(_("Customer paid amount must be greater than zero."))
        
        invoice_total = sum(invoice.amount_total for invoice in self.lead_id.invoice_ids if invoice.move_type == 'out_invoice' and invoice.payment_state == 'paid')

        if self.customer_paid > (self.rate - invoice_total):
            raise UserError(_("The customer paid amount exceeds the invoice total."))
        

        invoice = self.env['account.move'].create({
            'partner_id': self.partner_id.id,
            'move_type': 'out_invoice',
            'invoice_date': fields.Date.today(),
            'lead_id': self.lead_id.id,
            'invoice_line_ids': [(0, 0, {
                'product_id': self.property_product_id.id if self.property_product_id else False,
                'quantity': 1,
                'price_unit': self.customer_paid,
            })],
        })
        invoice.action_post()
        payment = self.env['account.payment'].create({
            'payment_type': 'inbound',
            'partner_type': 'customer',
            'partner_id': self.partner_id.id,
            'amount': self.customer_paid,
            'journal_id': self.env['account.journal'].search([('type', '=', 'bank')], limit=1).id,
            'payment_method_id': self.env.ref('account.account_payment_method_manual_in').id,
        })
        payment.action_post()
        invoice.payment_state = 'paid'
        if self.lead_id.rate == self.customer_paid + invoice_total:
            self.lead_id.invioce_fully_paid = True
            self.lead_id.payment_status = 'paid'
        else:
            self.lead_id.invioce_fully_paid = False
            self.lead_id.payment_status = 'partial'
        self.lead_id.customer_paid = self.customer_paid + invoice_total
        action = self.env["ir.actions.actions"]._for_xml_id("account.action_move_out_invoice_type")
        action['domain'] = [('id', '=', invoice.id)]
        action['context'] = {'form_view_initial_mode': 'edit'}
        
        return action