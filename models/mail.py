from odoo import models, fields, api, _
from xmlrpc import client as xmlrpclib
import email
import logging
import pytz
import re
from bs4 import BeautifulSoup
from datetime import datetime
import base64
import requests

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

class AccountMove(models.Model):
    _inherit = 'account.move'

    lead_id = fields.Many2one(
        'crm.lead',
        string='Lead',
        help="The lead associated with this invoice.",
    )



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

    @api.depends('rate', 'customer_paid')
    def _compute_balance(self):
        """Compute the balance amount to be paid by the customer."""
        for lead in self:
            if lead.rate and lead.customer_paid:
                lead.balance = lead.rate - lead.customer_paid
            else:
                lead.balance = 0.0

    # def _compute_today(self):
    #     """Compute the current date and time."""
    #     for lead in self:
    #         lead.today_start = datetime.now(pytz.timezone(self.env.user.tz or 'UTC')).replace(hour=0, minute=0, second=0, microsecond=0).astimezone(pytz.UTC).replace(tzinfo=None)
    #         lead.today_end = datetime.now(pytz.timezone(self.env.user.tz or 'UTC')).replace(hour=23, minute=59, second=59, microsecond=99).astimezone(pytz.UTC).replace(tzinfo=None)



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
            if email_from.endswith('agoda.com') or email_from == 'd365labs@gmail.com':

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
                    partner = self.env['res.partner'].create({
                        'name': f"{data.get('Customer First Name', '')} {data.get('Customer Last Name', '')}",
                        'email': data.get('Customer Email', ''),
                    })
                    in_date_obj = datetime.strptime(data.get('Check-in', ''), "%B %d, %Y").date()  # datetime.date(2025, 7, 9)
                    checkin = in_date_obj.strftime("%Y-%m-%d") 
                    out_date_obj = datetime.strptime(data.get('Check-out', ''), "%B %d, %Y").date()  # datetime.date(2025, 7, 9)
                    checkout = out_date_obj.strftime("%Y-%m-%d")
                    amount = float(data.get('Amount', '').replace(",", "").strip()) if data.get('Amount') else 0
                    net_rate = float(data.get('Net Rate', 0).replace(",", "").strip()) if data.get('Net Rate') else 0
                    lead = CRMLead.create({
                        'logo_src': logo_src,
                        'name': f"Agoda Booking {data.get('Booking ID', 'Unknown')} {data.get('Customer First Name', '')} {data.get('Customer Last Name', '')}",
                        'email_from': data.get('Customer Email', ''),
                        'city': data.get('City', ''),
                        'country_id': self.env['res.country'].search([('name', '=', data.get('Country of Residence', ''))], limit=1).id,
                        'check_in': checkin,
                        'check_out': checkout,
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
            if email_from.endswith('airbnb.com') or email_from == 'd365labs@gmail.com':
                if 'payout was sent' in msg_dict.get('subject', '').lower():
                    summary = {
                        "Airbnb Account ID": extract_field(r"Airbnb Account ID\s+(\d+)", cleaned_text),
                        "Payout ID": extract_field(r"\*Payout ID\s+([A-Za-z0-9]+)", cleaned_text),
                        "Payout Amount": extract_field(r"We've issued you a payout of\s+([₹\d,\.]+)", cleaned_text),
                        "Estimated Arrival": extract_field(r"account by\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})", cleaned_text),
                        "Total Amount Paid": extract_field(r"Amount paid.*?\n+([₹\d,\.]+)", cleaned_text)
                    }            
                    transactions = []
                    i = 0
                    while i < len(lines):
                        if lines[i] in ["Reservation", "Tax Withholding for India Income", "Home"]:
                            txn = {}
                            txn["type"] = lines[i]
                            txn["date_range"] = lines[i + 1]
                            txn["reservation_code"], guest_name, property_short = lines[i + 2].split(" - ", 2)
                            txn["guest_name"] = guest_name.strip()
                            txn["property_short"] = property_short.strip()

                            # Listing line
                            listing_line = lines[i + 3]
                            match = re.search(r"\(Listing ID:\s*(\d+)\)", listing_line)
                            if match:
                                txn["listing_id"] = match.group(1)
                            txn["property_full"] = listing_line.split(" (Listing ID")[0].strip()

                            raw_amount = lines[i + 4]
                            clean_amount = raw_amount.replace("₹", "").replace(",", "").strip()
                            txn["amount"] = float(clean_amount)

                            # Optional: parse check-in and check-out
                            checkin, checkout = txn["date_range"].split(" - ")
                            txn["check_in"] = checkin.strip()
                            txn["check_out"] = checkout.strip()

                            transactions.append(txn)
                            i += 5
                        else:
                            i += 1            
                    if summary.get('Airbnb Account ID'):
                        for transaction in transactions:
                            if transaction.get('type') in ['Reservation', 'Home']:
                                partner = self.env['res.partner'].create({
                                    'name': f"{transaction.get('guest_name', '')}",
                                    'email': transaction.get('email', ''),
                                })
                                lead = CRMLead.create({
                                    'logo_src': logo_src,
                                    'name': f"Airbnb Booking {transaction.get('reservation_code', 'Unknown')} {transaction.get('guest_name', '')}",
                                    'email_from': transaction.get('email', ''),
                                    'check_in': datetime.strptime(transaction.get('check_in', ''), "%m/%d/%Y").date(),
                                    'check_out': datetime.strptime(transaction.get('check_out', ''), "%m/%d/%Y").date(),
                                    'rate': transaction.get('amount', 0),
                                    'customer_paid': transaction.get('amount', 0),
                                    'partner_name': 'Airbnb',
                                    'partner_id': partner.id,
                                    'booking_id': transaction.get('reservation_code', ''),
                                    'net_rate': transaction.get('amount', 0),
                                    'payment_status': 'paid' if transaction.get('amount') else 'unpaid',
                                    'property_id': transaction.get('property_short', 0),
                                    'listing_id': transaction.get('listing_id', ''),
                                })
                                product = self.env['product.template'].search([('name', 'like', transaction.get('property_short'))], limit=1)

                                if product:
                                    lead.property_product_id = product.id
                                _logger.info('Created CRM Lead ID : %s', lead.id)
                                if transaction.get('amount') > 0:
                                    product = self.env['product.product'].search([('name', 'like', transaction.get('property_short', ''))], limit=1)
                                    invoice = self.env['account.move'].create({
                                        'partner_id': partner.id,
                                        'move_type': 'out_invoice',
                                        'invoice_date': datetime.now().date(),
                                        'lead_id': lead.id,
                                        'invoice_line_ids': [(0, 0, {
                                            'product_id': product.id if product else False,
                                            'quantity': 1,
                                            'price_unit': transaction.get('amount'),})],
                                    })
                                    invoice.action_post()
                                    payment = self.env['account.payment'].create({
                                        'payment_type': 'inbound',
                                        'partner_type': 'customer',
                                        'partner_id': partner.id,
                                        'amount': transaction.get('amount'),
                                        'journal_id': self.env['account.journal'].search([('type', '=', 'bank')], limit=1).id,
                                        'payment_method_id': self.env.ref('account.account_payment_method_manual_in').id,
                                    })
                                    payment.action_post()
                                    invoice.payment_state = 'paid'
                                    lead.invioce_fully_paid = True
                                    _logger.info('Created Invoice ID : %s', invoice.id)
                        _logger.info('Processed Airbnb booking with Account ID: %s', summary.get('Airbnb Account ID'))
                        return
            if email_from.endswith('go-mmt.com')  or email_from == 'd365labs@gmail.com':
                data = {
                    "Booking ID": extract_field(r"Booking ID\s+([A-Z0-9]+)", cleaned_text),
                    "Property Name": extract_field(r"Host Voucher \s+(.+?)", cleaned_text),
                    "City": extract_field(r"Yelahanka, (.+?)\n", cleaned_text),
                    "Customer First Name": extract_field(r"PRIMARY GUEST DETAILS\s+(.+?)\n", cleaned_text),
                    "Customer Last Name": "",  # not separately available, you can split first/last manually if needed
                    # Check-in and Check-out are below the headings, so find by scanning lines
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
                    partner = self.env['res.partner'].create({
                        'name': f"{data.get('Customer First Name', '')} {data.get('Customer Last Name', '')}",
                        'email': data.get('Customer Email', ''),
                    })
                    def parse_checkin_checkout(date_str):
                        try:
                            dt = datetime.strptime(date_str, "%d %b '%y %I:%M %p")
                            return dt  # return as naive datetime; Odoo handles server timezone conversion
                        except Exception as e:
                            _logger.error(f"Failed to parse date string: {date_str} — {e}")
                            return None
                    checkin = parse_checkin_checkout(data.get('Check-in', ''))
                    checkout = parse_checkin_checkout(data.get('Check-out', ''))
                    if checkin and checkout:
                        checkin = checkin.strftime("%Y-%m-%d %H:%M:%S")
                        checkout = checkout.strftime("%Y-%m-%d %H:%M:%S")
                    else:
                        checkin = checkout = None
                    
                    amount = float(data.get('Amount', '').replace(",", "").strip()) if data.get('Amount') else 0
                    net_rate = float(data.get('Net Rate', 0).replace(",", "").strip())
                    lead = CRMLead.create({
                        'logo_src': logo_src,
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

            if email_from.endswith('booking.com') or email_from == 'd365labs@gmail.com':

                links = soup.find_all("a", href=True)
                
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
                lead = CRMLead.create({
                        'logo_src': logo_src,
                        'name': f"Booking.com Booking {booking_id}",
                        'booking_url' : booking_url,
                        'partner_name': 'Booking.com',
                        'booking_id': booking_id,
                        'payment_status': 'unpaid',
                    })
                      
            
            
        # find possible routes for the message; note this also updates notably
        # 'author_id' of msg_dict
        routes = self.message_route(message, msg_dict, model, thread_id, custom_values)
        if self._detect_loop_sender(message, msg_dict, routes):
            return

        thread_id = self._message_route_process(message, msg_dict, routes)
        return thread_id