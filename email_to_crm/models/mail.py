from odoo import models, fields, api, _
from xmlrpc import client as xmlrpclib
import email
import logging
import pytz
import re
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import base64
import requests
from dateutil.parser import parse

# models/crm_lead.py
import io
from datetime import date
from openpyxl import Workbook
from odoo.exceptions import UserError

_logger = logging.getLogger(__name__)

class ProductTemplate(models.Model):
    _inherit = 'product.template'

    booking_dot_com_property_id = fields.Char(
        string='Booking.com Property ID',
        help="The Booking.com Property ID of this product.",
    )

    property_address = fields.Char(
        string='Property Address',
    )

    city = fields.Char(
        string='City',
        help="The city where the property is located.",
    )

    number_of_rooms = fields.Integer(
        string='Number of Rooms',
        help="The number of rooms.",
    )

    agoda_property_id = fields.Char(
        string='Agoda Property ID',
        help="The Agoda Property ID of this product.",
    )


    property_location = fields.Char(
        string='Property Location',
        help="The location of the property.",
    )

    make_my_trip_property_id = fields.Char(
        string='MakeMyTrip Property ID',
        help="The MakeMyTrip Property ID of this product.",
    )

    mrp_price = fields.Monetary(string="MRP")

    @api.model_create_multi
    def create(self, vals):
        res = super(ProductTemplate, self).create(vals)
        if not res.company_id and self.env.user.company_id:
            res.company_id = self.env.user.company_id.id
        return res



class CrmLead(models.Model):
    _inherit = 'crm.lead'

    property_id = fields.Char(
        string='Property',
        help="The property Booked by the customer.",
    )

    property_product_id = fields.Many2one(
        string='Property',
        comodel_name='product.template',
        help="The property Booked by the customer.",
    )

    logo_src = fields.Char(
        string='Logo Source',
        help="The source URL of the logo Booking Partner.",
    )

    logo = fields.Binary(
        string='Logo',
        compute='_compute_logo',
        store=True,
        help="The logo of the booking Partner.",
    )

    # today_start = fields.Datetime(
    #     string='Today',
    #     compute='_compute_today',
    #     store=False
    # )

    # today_end = fields.Datetime(
    #     string='Today', 
    #     compute='_compute_today', 
    #     store=False
    # )

    check_in = fields.Datetime(
        string='Check In',
        help="The check-in date and time of the customer.",
    )

    check_out = fields.Datetime(
        string='Check Out',
        help="The check-out date and time of the customer.",
    )

    booking_id = fields.Char(
        string='Booking ID',
        help="The booking ID from the booking partner.",
    )

    number_of_rooms = fields.Integer(
        string='Number of Rooms',
        help="The number of rooms Booked by the customer.",
    )

    rate = fields.Float(
        string='Total Rate',
        help="The rate of the property.",
    )

    customer_paid = fields.Float(
        string='Customer Paid',
        help="The amount paid by the customer.",
        compute='_compute_customer_paid',
    )

    balance = fields.Float(
        string='Balance',
        compute='_compute_balance',
        help="The balance amount to be paid by the customer.",
    )

    net_rate = fields.Float(
        string='Net Rate',
        help="The net rate promised to the customer.",
    )

    aadhar_id = fields.Binary(
        string='Aadhar ID',
        help="The Aadhar ID document of the customer.",
    )

    booking_url = fields.Char(
        string='Booking URL',
        help="The URL of the current booking - will redirect to the boking info page.",
        readonly=True
    )

    invoice_ids = fields.One2many(
        'account.move',
        'lead_id',
        string='Invoices',
        help="Invoices associated with this lead.",
    )

    invoice_count = fields.Integer(
        string='Invoice Count',
        compute='_compute_invoice_count',
        help="The number of invoices associated with this lead.",
    )

    invioce_fully_paid = fields.Boolean(
        string='Invoice Fully Paid',
    )

    city = fields.Char(
        string='City',
        help="The city where the lead is located.",
    )

    country_id = fields.Many2one(
        'res.country',
        string='Country',
        help="The country where the lead is located.",
    )

    listing_id = fields.Char(
        string='Listing ID',
        help="The listing ID associated with this lead.",
    )

    payment_status = fields.Selection(
        [('paid', 'Paid'), 
        ('unpaid', 'Unpaid'), 
        ('partial', 'Partially Paid')],
        string='Payment Status',
        help="The payment status of the lead.",
    )
    
    payment_transaction_id = fields.Text(
        string='Payment Transaction ID',
        help="The transaction ID of the payment made by the customer.",
    )

    other_guests = fields.Text(
        string='Other Guests',
        help="Information about other guests associated with this lead.",
    )

    payment_mode = fields.Char(
        string='Payment Mode',
        help="The mode of payment used by the customer.",
    )

    @api.model_create_multi
    def create(self, vals):
        res = super(CrmLead, self).create(vals)
        if not res.company_id and self.env.user.company_id:
            res.company_id = self.env.user.company_id.id
        return res

    @api.depends('rate', 'customer_paid')
    def _compute_balance(self):
        """Compute the balance amount to be paid by the customer."""
        for lead in self:
            if lead.rate and lead.customer_paid:
                lead.balance = lead.rate - lead.customer_paid
            else:
                lead.balance = 0.0

    @api.depends('invoice_ids')
    def _compute_customer_paid(self):
        """Compute the total amount paid by the customer based on related invoices."""
        for lead in self:
            total_paid = sum(invoice.amount_total for invoice in lead.invoice_ids if invoice.move_type == 'out_invoice' and invoice.payment_state in ['paid', 'in_payment'])
            lead.customer_paid = total_paid

    # @api.onchange('rate', 'payment_status')
    # def _onchange_rate_payment_status(self):
    #     """Update the customer_paid based on the payment_status."""
    #     if self.rate or self.payment_status:
    #         if self.payment_status == 'paid':
    #             self.customer_paid = self.rate
    #             self.invioce_fully_paid = True
    #         elif self.payment_status == 'unpaid':
    #             self.customer_paid = 0.0
    #             self.invioce_fully_paid = False
    #         elif self.payment_status == 'partial' and self.customer_paid > self.rate:
    #             self.customer_paid = self.rate
    #             self.invioce_fully_paid = False


    @api.depends('invoice_ids')
    def _compute_invoice_count(self):
        """Compute the number of invoices associated with this lead."""
        for lead in self:
            invoice_count = 0
            for invoice in lead.invoice_ids:
                if invoice.move_type == 'out_invoice':
                    invoice_count += 1
            lead.invoice_count = invoice_count

    def action_view_invoice(self):
        """Action to view the invoices associated with the lead."""
        self.ensure_one()
        action = self.env["ir.actions.actions"]._for_xml_id("account.action_move_out_invoice_type")
        action['domain'] = [('id', 'in', self.invoice_ids.ids),('move_type', '=', 'out_invoice')]
        action['context'] = {'form_view_initial_mode': 'edit'}  
        return action
    
    @api.depends('logo_src')
    def _compute_logo(self):
        """Compute the logo from the logo source URL."""
        for lead in self:
            if lead.logo_src:
                try:
                    response = requests.get(lead.logo_src)
                    if response.status_code == 200:
                        lead.logo = base64.b64encode(response.content)
                    else:
                        _logger.warning('Failed to fetch logo image from %s, status code: %s', lead.logo_src, response.status_code)
                        lead.logo = False
                except Exception as e:
                    _logger.error('Error fetching logo image from %s: %s', lead.logo_src, e)
                    lead.logo = False
            else:
                lead.logo = False

    @api.onchange('property_product_id')
    def _onchange_property_product_id(self):
        """Update the property_id and logo_src when the property_product_id changes."""
        if self.property_product_id:
            self.number_of_rooms = self.property_product_id.number_of_rooms
            self.city = self.property_product_id.city
        else:
            self.number_of_rooms = 0

    def create_invoice(self):
            inv_total = sum(invoice.amount_total for invoice in self.invoice_ids if invoice.move_type == 'out_invoice' and invoice.payment_state == 'paid')
            balance = self.rate - inv_total
            return {
            'type': 'ir.actions.act_window',
            'name': _('Create Invoice'),
            'res_model': 'create.invoice.wizard',
            'view_mode': 'form',
            'target': 'new',
            'context': {
                'default_partner_id': self.partner_id.id,
                'default_lead_id': self.id,
                'default_rate': self.rate,
                'default_customer_paid': balance if balance > 0 else 0,
                'default_property_product_id': self.property_product_id.id,
            },
        }
        
        
    @api.model
    def _generate_and_send_check_in_report(self, frequency):
        today = fields.Date.today()
        start = fields.Datetime.to_datetime(today)
        end = fields.Datetime.to_datetime(today) + timedelta(days=1)

        leads = self.search([('check_in', '>=', start),
                            ('check_in', '<', end),
                            ('company_id', '=', 1)])

        # Start HTML email body
        html_table = f"""
        <p>Hello,</p>
        <p>Please find below the <b>{frequency}</b> Check in Report for <b>{today.strftime('%d-%b-%Y')}</b>.</p>
        <table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse; width:100%; font-family:Arial; font-size:13px;">
            <thead style="background-color:#f2f2f2; text-align:center;">
                <tr>
                    <th>Guest Name</th>
                    <th>Room Name</th>
                    <th>City</th>
                    <th>Check-in (Today)</th>
<<<<<<< HEAD
                    <th>Payment Made (Today)</th>
                    <th>Payment Made (Total)</th>
=======
                    <th>Payment Made (Total Today)</th>
>>>>>>> 453fc7b307738f538317123eb3fbe620374e9d05
                    <th>Balance</th>
                    <th>Days Stay</th>
                </tr>
            </thead>
            <tbody>
        """

        total_payment_sum = 0
        total_balance_sum = 0

        for lead in leads:
            total_payment_today = 0

            # Calculate total payments from invoices
            for inv in lead.invoice_ids.filtered(lambda i: i.state == 'posted'):
                if frequency == 'Daily' and inv.invoice_date == today:
                    total_payment_today += inv.amount_total
                elif frequency == 'Weekly' and inv.invoice_date and inv.invoice_date >= today - timedelta(days=7):
                    total_payment_today += inv.amount_total



            total_payment_sum += total_payment_today
            total_balance_sum += lead.balance or 0

            # Calculate stay days
            days_stay = 0
            if lead.check_in and lead.check_out:
                days_stay = (lead.check_out - lead.check_in).days
            if lead.property_product_id.city:
                city = lead.property_product_id.city
            elif lead.city:
                city = lead.city
            else:
                city = ''

            html_table += f"""
                <tr>
                    <td>{lead.partner_id.name or ''}</td>
                    <td>{lead.property_product_id.name or ''}</td>
                    <td>{city or ''}</td>
                    <td>{lead.check_in.strftime('%d-%b-%Y') if lead.check_in else ''}</td>
                    <td style="text-align:right;">{total_payment_today:.2f}</td>
                    <td style="text-align:right;">{sum(inv.amount_total for inv in lead.invoice_ids.filtered(lambda i: i.state == 'posted')):.2f}</td>
                    <td style="text-align:right;">{lead.balance or 0:.2f}</td>
                    <td style="text-align:center;">{days_stay}</td>
                </tr>
            """

        # Add totals row
        html_table += f"""
            </tbody>
            <tfoot style="font-weight:bold; background-color:#e8e8e8;">
                <tr>
                    <td colspan="4" style="text-align:right;">Total:</td>
                    <td style="text-align:right;">{total_payment_sum:.2f}</td>
                    <td style="text-align:right;">{sum(inv.amount_total for inv in lead.invoice_ids.filtered(lambda i: i.state == 'posted')):.2f}</td>
                    <td style="text-align:right;">{total_balance_sum:.2f}</td>
                    <td></td>
                </tr>
            </tfoot>
        </table>
        <br>
        <p>Regards,<br/>Odoo System</p>
        """

        # Recipient email from system parameters
        recipient = self.env['ir.config_parameter'].sudo().get_param('crm.report_email')

        if recipient:
            mail_values = {
                'email_from': self.env.user.email_formatted,
                'subject': f'{frequency} CRM Report - {today.strftime("%d %B %Y")}',
                'body_html': html_table,
                'email_to': recipient,
            }
            self.env['mail.mail'].create(mail_values).send()
            _logger.info("CRM report email sent to %s", recipient)
        else:
            _logger.warning("No CRM report email configured or no check-in found for this date.")

    def _generate_daily_sales_report(self, frequency):
        today = date.today()
        today_paid_invoices = self.env['account.move'].search([
            ('move_type', '=', 'out_invoice'),
            ('invoice_date', '=', today),
            # ('payment_state', 'in', ['paid', 'in_payment','partial']),
            ('company_id', '=', 1)
        ])
        if not today_paid_invoices:
            return
        # # Prepare email content
        html_table = f"""
        <p>Hello,</p>
        <p>Please find below the <b>{frequency}</b> CRM Report for <b>{today.strftime('%d-%b-%Y')}</b>.</p>
        <table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse; width:100%; font-family:Arial; font-size:13px;">
            <thead style="background-color:#f2f2f2; text-align:center;">
                <tr>
                    <th>Room Name</th>
                    <th>Gust name</th>
                    <th>Price</th>
                    <th>Sold At</th>
                    <th>Payment Mode</th>
                    <th>Balance</th>
                    <th>City</th>
                </tr>
            </thead>
            <tbody>
        """
        total_payment_sum = 0
        total_balance_sum = 0
        for invoice in today_paid_invoices:
            if invoice.payment_state not in ['paid', 'in_payment','partial']:
                return
            lead = invoice.lead_id
            if not lead:
                continue
            total_payment = invoice.amount_total
            total_payment_sum += total_payment
            total_balance_sum += lead.balance or 0
            if lead.property_product_id.city:
                city = lead.property_product_id.city
            elif lead.city:
                city = lead.city
            else:
                city = ''
            html_table += f"""
                <tr>
                    <td>{lead.property_product_id.name or ''}</td>
                    <td>{lead.partner_id.name or ''}</td>
                    <td style="text-align:right;">{lead.property_product_id.list_price}</td>
                    <td>{invoice.amount_total:.2f}</td>
                    <td>{lead.payment_mode or ''}</td>
                    <td style="text-align:right;">{lead.balance or 0:.2f}</td>
                    <td>{city or ''}</td>
                </tr>
            """
        # Add totals row
        html_table += f"""
            </tbody>
            <tfoot style="font-weight:bold; background-color:#e8e8e8;">
                <tr>
                    <td style="text-align:right;" colspan="1">Total:</td>
                    <td></td>
                    <td></td>
                    <td style="text-align:right;">{total_payment_sum:.2f}</td>
                    <td></td>
                    <td style="text-align:right;">{total_balance_sum:.2f}</td>
                    <td></td>
                </tr>
            </tfoot>
        </table>
        <br>
        <p>Regards,<br/>Odoo System</p>
        """
        # Recipient email from system parameters
        recipient = self.env['ir.config_parameter'].sudo().get_param('crm.report_email')
        if recipient:
            mail_values = {
                'email_from': self.env.user.email_formatted,
                'subject': f'{frequency} Sales CRM Report - {today.strftime("%d %B %Y")}',
                'body_html': html_table,
                'email_to': recipient,
            }
            self.env['mail.mail'].create(mail_values).send()
            _logger.info("Sales CRM report email sent to %s", recipient)

    def _generate_daily_unsold_rooms_report(self):
        today = date.today()
        unsold_products = self.env['product.template'].search([
            ('company_id', '=', 1),
        ])
        unsold_rooms = []
        for product in unsold_products:
            bookings_count = self.search_count([
                ('property_product_id', '=', product.id),
                ('check_in', '>=', datetime.combine(today, datetime.min.time())),
                ('check_in', '<', datetime.combine(today + timedelta(days=1), datetime.min.time())),
            ])
            if bookings_count == 0:
                unsold_rooms.append(product)
        if not unsold_rooms:
            return
        # Prepare email content
        html_table = f"""
        <p>Hello,</p>
        <p>Please find below the Daily Unsold Rooms Report for <b>{today.strftime('%d-%b-%Y')}</b>.</p>
        <table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse; width:100%; font-family:Arial; font-size:13px;">
            <thead style="background-color:#f2f2f2; text-align:center;">
                <tr>
                    <th>Room Name</th>
                    <th>Price</th>
                    <th>City</th>
                    <th>Number of Rooms</th>
                </tr>
            </thead>
            <tbody>
        """
        for product in unsold_rooms:
            html_table += f"""
                <tr>
                    <td>{product.name or ''}</td>
                    <td>{product.list_price or ''}</td>
                    <td>{product.city or ''}</td>
                    <td style="text-align:right;">{product.number_of_rooms or 0}</td>
                </tr>
            """
        html_table += """
            </tbody>
        </table>
        <br>
        <p>Regards,<br/>Odoo System</p>
        """
        # Recipient email from system parameters
        recipient = self.env['ir.config_parameter'].sudo().get_param('crm.report_email')
        if recipient:
            mail_values = {
                'email_from': self.env.user.email_formatted,
                'subject': f'Daily Unsold Rooms Report - {today.strftime("%d %B %Y")}',
                'body_html': html_table,
                'email_to': recipient,
            }
            self.env['mail.mail'].create(mail_values).send()
            _logger.info("Unsold Rooms report email sent to %s", recipient)
        


class FetchmailServer(models.Model):
    """Incoming POP/IMAP mail server account"""

    _inherit = 'fetchmail.server'

    catch_mails_from = fields.Text(
        string='Catch Mails From',
        help="List of email addresses to catch emails from. "
             "If empty, all emails will be caught. "
             "You can use a comma-separated list of email addresses, "
             "This is useful to avoid catching emails from other users.",
        default='',
    )


class MailThread(models.AbstractModel):
    _inherit = 'mail.thread'

    @api.model
    def message_process(self, model, message, custom_values=None,
                        save_original=False, strip_attachments=False,
                        thread_id=None):
        """ Process an incoming RFC2822 email message, relying on
            ``mail.message.parse()`` for the parsing operation,
            and ``message_route()`` to figure out the target model.

            Once the target model is known, its ``message_new`` method
            is called with the new message (if the thread record did not exist)
            or its ``message_update`` method (if it did).

           :param string model: the fallback model to use if the message
               does not match any of the currently configured mail aliases
               (may be None if a matching alias is supposed to be present)
           :param message: source of the RFC2822 message
           :type message: string or xmlrpclib.Binary
           :type dict custom_values: optional dictionary of field values
                to pass to ``message_new`` if a new record needs to be created.
                Ignored if the thread record already exists, and also if a
                matching mail.alias was found (aliases define their own defaults)
           :param bool save_original: whether to keep a copy of the original
                email source attached to the message after it is imported.
           :param bool strip_attachments: whether to strip all attachments
                before processing the message, in order to save some space.
           :param int thread_id: optional ID of the record/thread from ``model``
               to which this mail should be attached. When provided, this
               overrides the automatic detection based on the message
               headers.
        """
        # extract message bytes - we are forced to pass the message as binary because
        # we don't know its encoding until we parse its headers and hence can't
        # convert it to utf-8 for transport between the mailgate script and here.
        if isinstance(message, xmlrpclib.Binary):
            message = bytes(message.data)
        if isinstance(message, str):
            message = message.encode('utf-8')
        message = email.message_from_bytes(message, policy=email.policy.SMTP)

        # parse the message, verify we are not in a loop by checking message_id is not duplicated
        msg_dict = self.message_parse(message, save_original=save_original)
        if strip_attachments:
            msg_dict.pop('attachments', None)

        existing_msg_ids = self.env['mail.message'].search([('message_id', '=', msg_dict['message_id'])], limit=1)
        if existing_msg_ids:
            _logger.info('Ignored mail from %s to %s with Message-Id %s: found duplicated Message-Id during processing',
                         msg_dict.get('email_from'), msg_dict.get('to'), msg_dict.get('message_id'))
            return False

        if self._detect_loop_headers(msg_dict):
            _logger.info('Ignored mail from %s to %s with Message-Id %s: reply to a bounce notification detected by headers',
                             msg_dict.get('email_from'), msg_dict.get('to'), msg_dict.get('message_id'))
            return
        fetch_list = []
        if self.env.context.get('params') or self.env.context.get('default_fetchmail_server_id') is not None:
            if self.env.context.get('params'):
                if self.env.context.get('params').get('model') == 'fetchmail.server':
                    emails_from_list = self.env['fetchmail.server'].browse(self.env.context.get('params').get('id')).catch_mails_from
                    if emails_from_list:
                        fetch_list = emails_from_list.split(',')
                    else:
                        fetch_list = []
            elif self.env.context.get('default_fetchmail_server_id'):
                emails_from_list = self.env['fetchmail.server'].browse(self.env.context.get('default_fetchmail_server_id')).catch_mails_from
                if emails_from_list:
                    fetch_list = emails_from_list.split(',')
                else:
                    fetch_list = []
        match = re.search(r'<([^>]+)>', msg_dict.get('email_from'))
        email_from = match.group(1)
        if fetch_list and email_from not in fetch_list:
            _logger.info('Ignored mail from %s to %s with Message-Id %s: email not in the catch list',
                         msg_dict.get('email_from'), msg_dict.get('to'), msg_dict.get('message_id'))
            return False
        else:
            _logger.info('Processing mail from %s to %s with Message-Id %s',
                         msg_dict.get('email_from'), msg_dict.get('to'), msg_dict.get('message_id'))
            CRMLead = self.env['crm.lead']
            soup = BeautifulSoup(msg_dict.get('body'), 'html.parser')
            img_tags = soup.find_all('img')
            if img_tags:
                for img in img_tags:
                    if 'src' in img.attrs and img['src'].startswith('https'):
                        logo_src = img.get('src')
                        break
                else:
                    logo_src = None
            else:
                logo_src = None
            # Extract text from the HTML content
            text = soup.get_text(separator="\n")
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            cleaned_text = '\n'.join(lines)
            def extract_field(pattern, text, default=None):
                    match = re.search(pattern, text)
                    return match.group(1).strip() if match else default
            if email_from.endswith('agoda.com') or email_from == 'd365labs@gmail.com' or email_from == 'sudarsanan1996@gmail.com':

                try:
                    idx = lines.index("Room Type")
                    room_type = lines[idx + 4] if len(lines) > idx + 4 else None
                    no_of_rooms = lines[idx + 5] if len(lines) > idx + 5 else None
                    occupancy = lines[idx + 6] if len(lines) > idx + 6 else None
                    extra_bed = lines[idx + 7] if len(lines) > idx + 7 else None
                except Exception as e:
                    room_type = no_of_rooms = occupancy = extra_bed = None
                

                data = {
                    "Booking ID": extract_field(r"Booking ID\s+(\d+)", cleaned_text),
                    "Property Name": extract_field(r"Booking confirmation\s+(.+?)\(", cleaned_text),
                    "Property ID": extract_field(r"Property ID\s*[\(:]?\s*(\d+)", cleaned_text),
                    "City": extract_field(r"City\s*:\s*(.+)", cleaned_text),
                    "Customer First Name": extract_field(r"Customer First Name\s+(.+)", cleaned_text),
                    "Customer Last Name": extract_field(r"Customer Last Name\s+(.+)", cleaned_text),
                    "Country of Residence": extract_field(r"Country of Residence\s+(.+)", cleaned_text),
                    "Check-in": extract_field(r"Check-in\s+(.+)", cleaned_text),
                    "Check-out": extract_field(r"Check-out\s+(.+)", cleaned_text),
                    "Other Guests": extract_field(r"Other Guests\s+(.+)", cleaned_text),
                    "Room Type": room_type,
                    "No. of Rooms": no_of_rooms,
                    "Occupancy": occupancy,
                    "Rate From-To": extract_field(r"From - To\s+Rates\s+([^\n]+)", cleaned_text),
                    "Amount": extract_field(r'INR\s*([\d,.]+)\s*\nReference sell rate', cleaned_text),
                    "Commission": extract_field(r'Commission\s*INR\s*(-?[\d,.]+)', cleaned_text),
                    "TDS": extract_field(r'TDS - Withholding tax\s*INR\s*(-?[\d,.]+)', cleaned_text),
                    "Rate Channel": extract_field(r"Rate Channel\s+(.+)", cleaned_text),
                    "Net Rate": extract_field(r"Net rate.*?INR\s*([\d,.]+)", cleaned_text),
                    "Customer Email": extract_field(r"Email:\s+(.+)", cleaned_text),
                    'payment_by': extract_field(r'Booked and Payable by\s*(.*?)\n', cleaned_text),
                }
                if data.get('Booking ID'):
                    if len(self.env['crm.lead'].search([('booking_id', '=', data.get('Booking ID'))])) == 0:

                        partner = self.env['res.partner'].create({
                            'name': f"{data.get('Customer First Name', '')} {data.get('Customer Last Name', '')}",
                            'email': data.get('Customer Email', ''),
                        })
                        user_tz = pytz.timezone(self.env.user.tz or 'Asia/Kolkata')
                        in_date_obj = datetime.strptime(data.get('Check-in', ''), "%B %d, %Y")
                        in_date_obj = user_tz.localize(in_date_obj.replace(hour=12, minute=0, second=0, microsecond=0)).astimezone(pytz.UTC).replace(tzinfo=None)
                        out_date_obj = datetime.strptime(data.get('Check-out', ''), "%B %d, %Y")
                        out_date_obj = user_tz.localize(out_date_obj.replace(hour=10, minute=0, second=0, microsecond=0)).astimezone(pytz.UTC).replace(tzinfo=None)
                        amount = float(data.get('Amount', '').replace(",", "").strip()) if data.get('Amount') else 0
                        net_rate = float(data.get('Net Rate', 0).replace(",", "").strip()) if data.get('Net Rate') else 0
                        lead = CRMLead.create({
                            'logo_src': 'email_to_crm/static/src/img/agoda.png' if not logo_src else logo_src,
                            'type': 'opportunity',
                            'name': f"Agoda Booking {data.get('Booking ID', 'Unknown')} {data.get('Customer First Name', '')} {data.get('Customer Last Name', '')}",
                            'email_from': data.get('Customer Email', ''),
                            'city': data.get('City', ''),
                            'country_id': self.env['res.country'].search([('name', '=', data.get('Country of Residence', ''))], limit=1).id,
                            'check_in': in_date_obj,
                            'check_out': out_date_obj,
                            'other_guests': data.get('Other Guests', ''),
                            'rate': amount,
                            'customer_paid': amount,
                            'partner_name': data.get('payment_by'),
                            'partner_id': partner.id,
                            'property_id': f"{data.get('Property Name', '')} ID: {data.get('Property ID', '')}",
                            'booking_id': data.get('Booking ID', ''),
                            'payment_status': 'paid' if data.get('Amount') else 'unpaid',
                            'net_rate': net_rate,
                        })
                        if data.get('Property ID'):
                            product = self.env['product.template'].search([('agoda_property_id', '=', data.get('Property ID'))], limit=1)
                            if product:
                                lead.property_product_id = product.id
                                lead.city = product.city
                        _logger.info('Created CRM Lead ID : %s', lead.id)
                        if amount > 0:
                            product = self.env['product.product'].search([('name', 'like', data.get('Property Name', ''))], limit=1)
                            invoice = self.env['account.move'].create({
                                'partner_id': partner.id,
                                'move_type': 'out_invoice',
                                'invoice_date': datetime.now().date(),
                                'lead_id': lead.id,
                                'invoice_line_ids': [(0, 0, {
                                    'product_id': product.id if product else False,
                                    'quantity': 1,
                                    'price_unit': amount,})],
                            })
                            invoice.action_post()
                            payment = self.env['account.payment'].create({
                                'payment_type': 'inbound',
                                'partner_type': 'customer',
                                'partner_id': partner.id,
                                'amount': amount,
                                'journal_id': self.env['account.journal'].search([('type', '=', 'bank')], limit=1).id,
                                'payment_method_id': self.env.ref('account.account_payment_method_manual_in').id,
                            })
                            payment.action_post()
                            invoice.payment_state = 'paid'
                            lead.invioce_fully_paid = True
                            _logger.info('Created Invoice ID : %s', invoice.id)
                        return
            if email_from.endswith('airbnb.com') or email_from == 'd365labs@gmail.com' or email_from == 'sudarsanan1996@gmail.com':
                if 'reservation confirmed' in msg_dict.get('subject', '').lower():
                    data = {}
                    name_match = re.search(r"New booking confirmed!\s*(.*?)\s*arrives", cleaned_text)
                    if name_match:
                        data['guest_name'] = name_match.group(1).strip() 
                    # Check-in date
                    checkin_match = re.search(r"Check-in\s*([A-Za-z]+,\s*[A-Za-z]+\s*\d+)", text)
                    if checkin_match:
                        checkin_match = parse(checkin_match.group(1).strip())
                        checkin_match = checkin_match.replace(year=datetime.today().year)
                        data['checkin_date'] = checkin_match

                    # Check-in time
                    checkin_time = re.search(r"Check-in.*?(\d{1,2}:\d{2}\s*[APM]{2})", text)
                    if checkin_time:
                        data['checkin_time'] = checkin_time.group(1)

                    # Checkout date
                    checkout_match = re.search(r"Checkout\s*([A-Za-z]+,\s*[A-Za-z]+\s*\d+)", text)
                    if checkout_match:
                        checkout_match = parse(checkout_match.group(1).strip())
                        checkout_match = checkout_match.replace(year=datetime.today().year)
                        data['checkout_date'] = checkout_match

                    # Checkout time
                    checkout_time = re.search(r"Checkout.*?(\d{1,2}:\d{2}\s*[APM]{2})", text)
                    if checkout_time:
                        data['checkout_time'] = checkout_time.group(1)
                        t = parse(checkout_time.group(1)).time()
                        data['checkout_date'] = data['checkout_date'].replace(
                            hour=t.hour, minute=t.minute, second=0, microsecond=0
                        )
                    else:
                        # Default Airbnb checkout
                        data['checkout_date'] = data['checkout_date'].replace(
                            hour=10, minute=0, second=0
                        )

                    if data.get('checkin_time'):
                        t = parse(data['checkin_time']).time()
                        data['checkin_date'] = data['checkin_date'].replace(
                            hour=t.hour, minute=t.minute, second=0, microsecond=0
                        )
                    else:
                        # Default Airbnb checkin
                        data['checkin_date'] = data['checkin_date'].replace(
                            hour=12, minute=0, second=0
                        )
                    

                    # ===========================================
                    # Convert to UTC naive datetime for Odoo
                    # ===========================================
                    tz = pytz.timezone('Asia/Kolkata')
                    data['check_in'] = tz.localize(data['checkin_date']).astimezone(pytz.UTC).replace(tzinfo=None)
                    data['check_out'] = tz.localize(data['checkout_date']).astimezone(pytz.UTC).replace(tzinfo=None)

                    # Guests
                    guests = re.search(r"Guests\s*([\d]+\s*adults?,\s*[\d]+\s*children?)", text)
                    if guests:
                        data['guest_count'] = guests.group(1)

                    # Confirmation code
                    confirmation = re.search(r"Confirmation code\s*([A-Z0-9]+)", text)
                    if confirmation:
                        data['confirmation_code'] = confirmation.group(1)

                    # Guest paid total
                    guest_total = re.search(r"Total \(INR\)\s*₹([\d,]+\.\d+)", text)
                    if guest_total:
                        data['guest_total'] = float(guest_total.group(1).replace(',', ''))

                    # Host earns
                    host_earn = re.search(r"You earn\s*₹([\d,]+\.\d+)", text)
                    if host_earn:
                        data['host_earnings'] = host_earn.group(1)

                    # Occupancy taxes
                    tax_match = re.search(r"Occupancy taxes\s*₹([\d,]+\.\d+)", text)
                    if tax_match:
                        data['tax_amount'] = tax_match.group(1)
                    property_match = re.search(r"([\w\s\d,.-]+?)\s*Entire home/apt", text)
                    if property_match:
                        data['property_name'] = property_match.group(1).strip()
     
                    if len(self.env['crm.lead'].search([('booking_id', '=', data.get('confirmation_code'))])) == 0:
                        partner = self.env['res.partner'].create({
                            'name': f"{data.get('guest_name', '')}",
                            'email': data.get('email', ''),
                        })
                        lead = CRMLead.create({
                            'logo_src': 'email_to_crm/static/src/img/Airbnb_Logo.png' if not logo_src else logo_src,
                            'type': 'opportunity',
                            'name': f"Airbnb Booking {data.get('confirmation_code', 'Unknown')} {data.get('guest_name', '')}",
                            'email_from': data.get('email', ''),
                            'check_in': data.get('check_in', ''),
                            'check_out':data.get('check_out', ''),
                            'rate': data.get('guest_total', 0),
                            'customer_paid': data.get('guest_total', 0),
                            'partner_name': 'Airbnb',
                            'partner_id': partner.id,
                            'booking_id': data.get('confirmation_code', ''),
                            'net_rate': data.get('guest_total', 0),
                            'payment_status': 'paid' if data.get('guest_total') else 'unpaid',
                            'property_id': data.get('property_name', 0),
                        })
                        product = self.env['product.template'].search([('name', 'like', data.get('property_name'))], limit=1)

                        if product:
                            lead.property_product_id = product.id
                            lead.city = product.city
                        _logger.info('Created CRM Lead ID : %s', lead.id)
                        if data.get('guest_total') > 0:
                            product = self.env['product.product'].search([('name', 'like', data.get('property_name', ''))], limit=1)
                            invoice = self.env['account.move'].create({
                                'partner_id': partner.id,
                                'move_type': 'out_invoice',
                                'invoice_date': datetime.now().date(),
                                'lead_id': lead.id,
                                'invoice_line_ids': [(0, 0, {
                                    'product_id': product.id if product else False,
                                    'quantity': 1,
                                    'price_unit': data.get('guest_total'),})],
                            })
                            invoice.action_post()
                            payment = self.env['account.payment'].create({
                                'payment_type': 'inbound',
                                'partner_type': 'customer',
                                'partner_id': partner.id,
                                'amount': data.get('guest_total'),
                                'journal_id': self.env['account.journal'].search([('type', '=', 'bank')], limit=1).id,
                                'payment_method_id': self.env.ref('account.account_payment_method_manual_in').id,
                            })
                            payment.action_post()
                            invoice.payment_state = 'paid'
                            lead.invioce_fully_paid = True
                            _logger.info('Created Invoice ID : %s', invoice.id)
                        _logger.info('Processed Airbnb booking for : %s', data.get('guest_name'))
                        return
            if email_from.endswith('go-mmt.com')  or email_from == 'd365labs@gmail.com' or email_from == 'sudarsanan1996@gmail.com':
                data = {
                    "Booking ID": extract_field(r"Booking ID\s+([A-Z0-9]+)", cleaned_text),
                    "Property Name": extract_field(r"Host Voucher \s+(.+?)", cleaned_text),
                    "City": extract_field(r"Yelahanka, (.+?)\n", cleaned_text),
                    "Customer First Name": extract_field(r"PRIMARY GUEST DETAILS\s+(.+?)\n", cleaned_text),
                    "Customer Last Name": "",  # not separately available, you can split first/last manually if needed
                    "Check-in": next((lines[i + 2] + " "+ lines[i + 3] for i, line in enumerate(lines) if line.strip().upper() == "CHECK-IN" and i + 1 < len(lines)), None),
                    "Check-out": next((lines[i + 3] + " " + lines[i + 5] for i, line in enumerate(lines) if line.strip().upper() == "CHECK-OUT" and i + 1 < len(lines)), None),
                    "No. of Rooms": extract_field(r"Room\(s\)\s+(\d+)", cleaned_text),
                    "Room Type": extract_field(r"x (.+?)\n", cleaned_text),
                    "Occupancy": extract_field(r"TOTAL NO\. OF GUEST\(S\)\s+(.+)", cleaned_text),
                    "Amount": extract_field(r"Property Gross Charges\s+₹\s*([\d,.]+)", cleaned_text),
                    "Commission": extract_field(r"Go-MMT Commission\s+₹\s*([\d,.]+)", cleaned_text),
                    "TDS": extract_field(r"TDS @ [\d.]+%\s+₹\s*([\d,.]+)", cleaned_text),
                    "Net Rate": extract_field(r"Payable to Property\s+₹\s*([\d,.]+)", cleaned_text),
                    "Rate Channel": "MakeMyTrip",
                    "Customer Email": "",  # Not available in text
                    "payment_by": extract_field(r"Payment Status\s+(.+)", cleaned_text),
                }
                if data.get('Booking ID'):
                    if len(self.env['crm.lead'].search([('booking_id', '=', data.get('Booking ID'))])) == 0:
                        partner = self.env['res.partner'].create({
                            'name': f"{data.get('Customer First Name', '')} {data.get('Customer Last Name', '')}",
                            'email': data.get('Customer Email', ''),
                        })
                        def parse_checkin_checkout(date_str):
                            try:
                                # Extract only the part that looks like "02 Oct '25" or "30 Sep '25 12:00 PM"
                                match = re.search(r"\d{2} \w{3} '\d{2}(?: \d{1,2}:\d{2} (AM|PM))?", date_str)
                                if not match:
                                    raise ValueError("No valid date pattern found")
                                
                                clean_date = match.group(0)

                                # Try parsing with datetime+time
                                try:
                                    dt = datetime.strptime(clean_date, "%d %b '%y %I:%M %p")
                                except ValueError:
                                    # Fall back to date only
                                    dt = datetime.strptime(clean_date, "%d %b '%y")

                                return dt  # naive datetime (Odoo handles TZ)
                            
                            except Exception as e:
                                _logger.error(f"Failed to parse date string: {date_str} — {e}")
                                return None
                        checkin = parse_checkin_checkout(data.get('Check-in', ''))
                        checkout = parse_checkin_checkout(data.get('Check-out', ''))
                        if checkin and checkout:
                            user_tz = pytz.timezone(self.env.user.tz or 'Asia/Kolkata')
                            checkin = user_tz.localize(checkin.replace(hour=12, minute=0, second=0, microsecond=0)).astimezone(pytz.UTC).replace(tzinfo=None)
                            checkout = user_tz.localize(checkout.replace(hour=10, minute=0, second=0, microsecond=0)).astimezone(pytz.UTC).replace(tzinfo=None)
                        else:
                            checkin = checkout = None
                        
                        amount = float(data.get('Amount', '').replace(",", "").strip()) if data.get('Amount') else 0
                        net_rate = float(data.get('Net Rate', 0).replace(",", "").strip())
                        lead = CRMLead.create({
                            'logo_src': 'email_to_crm/static/src/img/mmt.png' if not logo_src else logo_src,
                            'type': 'opportunity',
                            'name': f"MakeMyTrip Booking {data.get('Booking ID', '')} {data.get('Customer First Name', '')}",
                            'check_in': checkin,
                            'check_out': checkout,
                            'rate': amount,
                            'customer_paid': amount,
                            'number_of_rooms' : int(data.get('No. of Rooms', 0)) if data.get('No. of Rooms') else 0,
                            'partner_name': 'MakeMyTrip',
                            'partner_id': partner.id,
                            'booking_id': data.get('Booking ID', ''),
                            'net_rate': net_rate,
                            'payment_status': 'paid' if amount else 'unpaid',
                            'property_id': data.get('Room Type', 0),
                        })
                        product = self.env['product.template'].search([('name', 'like', data.get('Room Type'))], limit=1)
                        if product:
                            lead.property_product_id = product.id
                            lead.city = product.city
                        _logger.info('Created CRM Lead ID : %s', lead.id)
                        if amount > 0:
                            product = self.env['product.product'].search([('name', 'like', data.get('Room Type', ''))], limit=1)
                            invoice = self.env['account.move'].create({
                                'partner_id': partner.id,
                                'move_type': 'out_invoice',
                                'invoice_date': datetime.now().date(),
                                'lead_id': lead.id,
                                'invoice_line_ids': [(0, 0, {
                                    'product_id': product.id if product else False,
                                    'quantity': 1,
                                    'price_unit': amount,})],
                            })
                            invoice.action_post()
                            payment = self.env['account.payment'].create({
                                'payment_type': 'inbound',
                                'partner_type': 'customer',
                                'partner_id': partner.id,
                                'amount': amount,
                                'journal_id': self.env['account.journal'].search([('type', '=', 'bank')], limit=1).id,
                                'payment_method_id': self.env.ref('account.account_payment_method_manual_in').id,
                            })
                            payment.action_post()
                            # Link the payment with the invoice
                            invoice.payment_state = 'paid'
                            lead.invioce_fully_paid = True
                            _logger.info('Created Invoice ID : %s', invoice.id)
                        return
            if email_from.endswith('booking.com') or email_from == 'd365labs@gmail.com' or email_from == 'sudarsanan1996@gmail.com':

                links = soup.find_all("a", href=True)
                booking_node = soup.find(text=re.compile("Booking.com"))
                property_name = None
                if booking_node:
                    # Get the next text after Booking.com
                    next_text = booking_node.find_next(string=True)
                    if next_text:
                        property_name = next_text.strip()
                
                booking_data = None
                # 2. Filter for booking.com URLs containing res_id
                for link in links:
                    if link.text:
                        href = link.text.strip()
                        if "admin.booking.com" in href and "res_id=" in href:
                            # Extract booking ID from query params using regex
                            match = re.search(r"res_id=(\d+)", href)
                            booking_id = match.group(1) if match else None
                            booking_data = {
                                "url": href,
                                "booking_id": booking_id
                            }
                            break
                        else:
                            continue
                    else:
                        continue
                

                # Step 1: URL and credentials
                if not booking_data:
                    _logger.warning('No valid booking.com link found in the email.')
                    return
                booking_url = booking_data['url']
                booking_id = booking_data['booking_id']
                if property_name:
                    property_search = self.env['product.product'].search([('name', 'ilike', property_name)], limit=1)
                    property_id = property_search.id if property_search else None
                else:
                    property_id = None
                if len(self.env['crm.lead'].search([('booking_id', '=', booking_id)])) == 0:
                    lead = CRMLead.create({
                        'logo_src': logo_src,
                        'type': 'opportunity',
                        'name': f"Booking.com Booking {booking_id}",
                        'booking_url' : booking_url,
                        'partner_name': 'Booking.com',
                        'booking_id': booking_id,
                        'property_product_id': property_id,
                        'payment_status': 'unpaid',
                    })
                      
            
            
        # find possible routes for the message; note this also updates notably
        # 'author_id' of msg_dict
        routes = self.message_route(message, msg_dict, model, thread_id, custom_values)
        if self._detect_loop_sender(message, msg_dict, routes):
            return

        thread_id = self._message_route_process(message, msg_dict, routes)
        return thread_id
