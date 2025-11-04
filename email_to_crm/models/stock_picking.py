from odoo import models, fields, api
from openpyxl import Workbook
import io, base64
import logging
from datetime import date, timedelta

_logger = logging.getLogger(__name__)

class StockPicking(models.Model):
    _inherit = 'stock.picking'

    @api.model_create_multi
    def create(self, vals):
        sales_man = self.env['stock.location'].search([('name', '=', 'SALES MAN')], limit=1)
        for val in vals:
            if sales_man and not val.get('location_id'):
                val['location_id'] = sales_man.id
        return super().create(vals)
    
    def _send_daily_return_report(self):
        """Send daily return report to configured CRM email."""
        # crm_email = self.env['ir.config_parameter'].sudo().get_param('account.report_email')
        # if not crm_email:
        #     return

        today = fields.Date.today()
        return_pickings = self.search([
            ('picking_type_id.code', '=', 'return'),
            ('scheduled_date', '>=', today),
            ('scheduled_date', '<', today + timedelta(days=1)),
            ('state', '=', 'done')
        ])

        if not return_pickings:
            return

        # Prepare email content
        wb = Workbook()
        ws = wb.active
        ws.append(['Return Reference', 'Customer', 'Date', 'Total Quantity'])
        for picking in return_pickings:
            ws.append([
                picking.name,
                picking.partner_id.name,
                picking.scheduled_date,
                sum(move.product_uom_qty for move in picking.move_lines)
            ])
        # Save the workbook to a binary stream
        fp = io.BytesIO()
        if ws.max_row > 1:
            wb.save(fp)
        else:
            fp.close()
            _logger.info('No data rows after headers; Excel report not generated.')
            return
        fp.seek(0)
        file_data = base64.b64encode(fp.read())
        fp.close()
        # Create attachment
        attachment = self.env['ir.attachment'].create({
            'name': f"Daily_Return_Report_{today}.xlsx",
            'type': 'binary',
            'datas': file_data,
            'res_model': 'stock.picking',
            'res_id': 0,
        })
