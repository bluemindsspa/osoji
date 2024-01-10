from odoo import _, api, fields, models
import base64
import xlwt
from io import BytesIO
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
import logging

_logger = logging.getLogger(__name__)

class AccountBookReport(models.TransientModel):
    _name = 'account.book.report'
    _description = 'Account Book Report'

    def default_company_user(self):
        return self.env.user.company_id.id

    name = fields.Char(string='Detalle')
    #type_operation = fields.Selection([('sell', 'Venta'), ('buy', 'Compra'), ('ticket', 'Boleta Electrónica'), ('fee', 'Honorarios')],
    type_operation = fields.Selection([('sell', 'Venta'), ('buy', 'Compra'), ('ticket', 'Boleta Electrónica')],
        "Tipo de operación", required=True, default="sell")
    date = fields.Date("Fecha", default=lambda self: fields.Date.context_today(self))
    tax_period = fields.Char('Periodo Tributario', required=True, default=lambda *a: datetime.now().strftime('%Y-%m'))
    invoice_ids = fields.Many2many('account.move', string='Documentos')
    company_id = fields.Many2one('res.company', string='Compañía', default=default_company_user)

    @api.onchange('tax_period', 'type_operation')
    def onchange_tickets_by_operation(self):
        cr = self._cr
        current = datetime.strptime(self.tax_period + '-01', '%Y-%m-%d')
        next_month = current + relativedelta(months=1)
        lines = []
        if self.type_operation in ['sell', 'ticket']:
            # query = "SELECT id FROM account_move WHERE date >= '%s' and date < '%s' and move_type in ('out_invoice', 'out_refund') and state = 'posted' and company_id = %s ORDER BY id ASC;" % (current, next_month, self.company_id.id)
            query = "SELECT a.id, t.code FROM account_move a " \
                    "INNER JOIN l10n_latam_document_type t ON t.id = a.l10n_latam_document_type_id " \
                    "WHERE a.date >= '%s' and a.date < '%s' and a.move_type in ('out_invoice', 'out_refund') and a.state = 'posted' and a.company_id = %s " \
                    "ORDER BY t.code ASC;" % (current, next_month, self.company_id.id)
        else:
            # query = "SELECT id FROM account_move WHERE date >= '%s' and date < '%s' and move_type in ('in_invoice', 'in_refund') and state = 'posted' and company_id = %s ORDER BY id ASC;" % (current, next_month, self.company_id.id)
            query = "SELECT a.id, t.code FROM account_move a " \
                    "INNER JOIN l10n_latam_document_type t ON t.id = a.l10n_latam_document_type_id " \
                    "WHERE a.date >= '%s' and a.date < '%s' and a.move_type in ('in_invoice', 'in_refund') and a.state = 'posted' and a.company_id = %s " \
                    "ORDER BY t.code ASC;" % (current, next_month, self.company_id.id)
        cr.execute(query)
        for row in cr.dictfetchall():
            lines.append(row['id'])
        self.invoice_ids = [(6, 0, lines)]

    def report_name(self, name):
        if self.type_operation == 'ticket':
            return 'Boleta ' + self.tax_period
        elif self.type_operation == 'sell':
            return 'Venta ' + self.tax_period
        else:
            return 'Compra ' + self.tax_period

    def _get_totals_invoices(self, lines):
        amount = exempt = amount_iva = other_imps = amount_12 = 0
        for l in lines:
            if l.tax_ids:
                amount += l.price_subtotal
                for tax in l.tax_ids:
                    if tax.l10n_cl_sii_code == 0:
                        exempt += l.price_subtotal
                        amount -= exempt
                    elif tax.l10n_cl_sii_code == 14:
                        amount_iva += l.price_subtotal * tax.amount / 100

                    elif tax.l10n_cl_sii_code == 19:
                        amount_12 += l.price_subtotal * tax.amount / 100
                    else:
                        other_imps += l.price_subtotal * tax.amount / 100
            else:
                exempt += l.price_subtotal
        totals = {
            'amount': amount,
            'exempt': exempt,
            'iva': amount_iva,
            'other_imps': other_imps,
            'amount_12': amount_12
        }
        return totals

    def _get_totals_invoices_currency(self, lines):
        amount = exempt = amount_iva = other_imps = amount_12 = 0
        for l in lines:
            if l.tax_ids:
                amount += l.credit if l.credit else l.debit
                for tax in l.tax_ids:
                    if tax.l10n_cl_sii_code == 0:
                        exempt += l.credit if l.credit else l.debit
                        amount -= exempt
                    elif tax.l10n_cl_sii_code == 14:
                        amount_iva += (l.credit if l.credit else l.debit) * tax.amount / 100

                    elif tax.l10n_cl_sii_code == 19:
                        amount_12 += (l.credit if l.credit else l.debit) * tax.amount / 100
                    else:
                        other_imps += (l.credit if l.credit else l.debit) * tax.amount / 100
            else:
                exempt += l.credit if l.credit else l.debit
        totals = {
            'amount': amount,
            'exempt': exempt,
            'iva': amount_iva,
            'other_imps': other_imps,
            'amount_12': amount_12
        }
        return totals

    def export_report_xls(self):
        report_name = self.report_name(self.name)
        company = self.env.user.company_id
        workbook = xlwt.Workbook(encoding="utf-8")
        sheet = workbook.add_sheet(report_name)
        bold = xlwt.easyxf('font: bold on;')
        sheet.write(0, 0, company.name, bold)
        sheet.write(0, 1, 'Fecha: ' + datetime.strftime(self.date, '%d/%m/%Y'))
        sheet.write(1, 0, company.l10n_cl_activity_description, bold)
        sheet.write(2, 0, company.partner_id.vat, bold)
        sheet.write(3, 0, report_name, bold)
        sheet.write(4, 0, self.tax_period, bold)
        sheet.write(5, 0, self.type_operation, bold)
        sheet.write(9, 0, "Tipo de Documento", bold)
        sheet.write(9, 1, "Número", bold)
        sheet.write(9, 2, "Fecha Emisión", bold)
        sheet.write(9, 3, "RUT", bold)
        sheet.write(9, 4, "Entidad", bold)
        sheet.write(9, 5, "Afecto", bold)
        sheet.write(9, 6, "Exento", bold)
        sheet.write(9, 7, "IVA", bold)
        #if self.type_operation in ['sell', 'fee']:
        if self.type_operation in ['sell']:
            sheet.write(9, 8, "ANT 12%IVA", bold)
            sheet.write(9, 9, "Total", bold)
        else:
            sheet.write(9, 8, "Total", bold)
        line = 10
        ########################### DETAILS ##########################################
        total_details = exempt_details = amount_details = tax_details = total_12_details = other_imps = 0
        moves = []
        if self.type_operation == 'ticket':
            invoices = self.invoice_ids.filtered(lambda s: s.l10n_latam_document_type_id.code in ['39'] and s.journal_id.type == 'sale')
            for inv in invoices:
                moves.append(inv)
        else:
            invoices = self.invoice_ids.filtered(lambda s: s.l10n_latam_document_type_id.code in ['33', '34', '39','56', '61', '110', '112'] and s.journal_id.type == 'sale') if self.type_operation == 'sell' \
            else self.invoice_ids.filtered(lambda s: s.l10n_latam_document_type_id.code in ['33', '34', '56', '61', '914'] and s.journal_id.type == 'purchase') if self.type_operation == 'buy' \
            else self.invoice_ids.filtered(lambda s: s.l10n_latam_document_type_id.code in ['71'] and s.journal_id.type == 'purchase')
            document_class = invoices.mapped("l10n_latam_document_type_id")
            for inv in invoices:
                moves.append(inv)
        for move in moves:
            _logger.info('Factura #%s' % move.name)
            sheet.write(line, 0, move.l10n_latam_document_type_id.name)
            sheet.write(line, 1, move.l10n_latam_document_number)
            sheet.write(line, 2, datetime.strftime(move.invoice_date, '%d/%m/%Y'))
            sheet.write(line, 3, move.partner_id.vat)
            sheet.write(line, 4, move.partner_id.name)
            if move.l10n_latam_document_type_id.code in ['110', '112']:
                totals = self._get_totals_invoices_currency(move.invoice_line_ids)
            else:
                totals = self._get_totals_invoices(move.invoice_line_ids)
            sheet.write(line, 5, totals['amount'])
            sheet.write(line, 6, totals['exempt'])
            sheet.write(line, 7, totals['iva'])
            if self.type_operation in ['sell', 'fee']:
                total_12_details += totals['amount_12']
                sheet.write(line, 8, totals['amount_12'])
                sheet.write(line, 9, abs(move.amount_total_signed))
            else:
                sheet.write(line, 8, abs(move.amount_total_signed))
            if move.l10n_latam_document_type_id.code in ['61', '112']:
                amount_details -= totals['amount']
                exempt_details -= totals['exempt']
                tax_details -= move.amount_tax
                other_imps -= totals['other_imps']
                total_details -= abs(move.amount_total_signed)
            else:
                amount_details += totals['amount']
                exempt_details += totals['exempt']
                tax_details += move.amount_tax
                other_imps += totals['other_imps']
                total_details += abs(move.amount_total_signed)
            line += 1
        sheet.write(line, 0, "Total General", bold)
        sheet.write(line, 5, amount_details, bold)
        sheet.write(line, 6, exempt_details, bold)
        sheet.write(line, 7, tax_details, bold)
        if self.type_operation in ['sell', 'fee']:
            sheet.write(line, 8, total_12_details, bold)
            sheet.write(line + 1, 0, 'Total (Afecto + Excento + Iva + Otros Imp)', bold)
            sheet.write(line + 1, 9, amount_details + exempt_details + tax_details + other_imps, bold)
        else:
            sheet.write(line + 1, 0, 'Total (Afecto + Excento + Iva + Otros Imp)', bold)
            sheet.write(line + 1, 8, amount_details + exempt_details + tax_details + other_imps, bold)
        ############################### RESUME ############################################
        total_documents_general = 0
        amount_global = exempt_global = tax_global = total_12_global = other_imps = 0
        line += 5
        sheet.write(line, 0, "Tipo de Documento", bold)
        sheet.write(line, 1, "Cantidad de Documentos", bold)
        sheet.write(line, 2, "Afecto", bold)
        sheet.write(line, 3, "Exento", bold)
        sheet.write(line, 4, "IVA", bold)
        if self.type_operation in ['sell', 'fee']:
            sheet.write(line, 5, "ANT 12%IVA", bold)
            sheet.write(line, 6, "Total", bold)
        else:
            sheet.write(line, 5, "Total", bold)
        if self.type_operation == 'ticket':
            invoices = self.invoice_ids.filtered(lambda s: s.l10n_latam_document_type_id.code in ['39'])
            document_class = invoices.mapped("l10n_latam_document_type_id")
            for inv in invoices:
                moves.append(inv)
        else:
            invoices = self.invoice_ids.filtered(lambda s: s.l10n_latam_document_type_id.code in ['33', '34', '39','56', '61', '110', '112'] and s.journal_id.type == 'sale') if self.type_operation == 'sell' \
            else self.invoice_ids.filtered(lambda s: s.l10n_latam_document_type_id.code in ['33', '34', '56', '61', '914'] and s.journal_id.type == 'purchase') if self.type_operation == 'buy' \
            else self.invoice_ids.filtered(lambda s: s.journal_id.type == 'purchase')
            document_class = invoices.mapped("l10n_latam_document_type_id")
            for inv in invoices:
                moves.append(inv)
        document_class = list(set(document_class))
        for dc in document_class:
            total_documents = []
            total_general = exempt_general = amount_general = tax_general = total_12_general = 0
            if self.type_operation == 'ticket':
                documents = invoices.filtered(lambda s: s.l10n_latam_document_type_id.name == dc.name)
                for doc in documents:
                    total_documents.append(doc)
            else:
                documents = invoices.filtered(lambda s: s.l10n_latam_document_type_id.name == dc.name)
                for doc in documents:
                    total_documents.append(doc)
            sheet.write(line + 1, 0, dc.name)
            sheet.write(line + 1, 1, len(total_documents))
            total_documents_general += len(total_documents)
            if self.type_operation == 'ticket':
                for doc in total_documents:
                    if doc._name == 'account.move':
                        totals = self._get_totals_invoices(move.invoice_line_ids)
                        amount_general += totals['amount']
                        exempt_general += totals['exempt']
                        other_imps += totals['other_imps']
                        tax_general += doc.amount_tax
                        total_general += doc.amount_total
                sheet.write(line + 1, 2, amount_general)
                sheet.write(line + 1, 3, exempt_general)
                amount_global += amount_general
                exempt_global += exempt_general
                tax_global += tax_general
            elif self.type_operation in ['sell', 'fee']:
                for doc in total_documents:
                    if doc.l10n_latam_document_type_id.code in ['110', '112']:
                        totals = self._get_totals_invoices_currency(doc.invoice_line_ids)
                    else:
                        totals = self._get_totals_invoices(doc.invoice_line_ids)
                    if dc.code in ['61', '112']:
                        amount_general -= totals['amount']
                        exempt_general -= totals['exempt']
                        tax_general -= doc.amount_tax
                        total_12_general -= totals['amount_12']
                        total_general -= abs(doc.amount_total_signed)
                    else:
                        amount_general += totals['amount']
                        exempt_general += totals['exempt']
                        tax_general += doc.amount_tax
                        total_12_general += totals['amount_12']
                        total_general += abs(doc.amount_total_signed)
                sheet.write(line + 1, 2, abs(amount_general))
                sheet.write(line + 1, 3, abs(exempt_general))
                amount_global += amount_general
                exempt_global += exempt_general
                tax_global += tax_general
            else:
                for doc in total_documents:
                    totals = self._get_totals_invoices(doc.invoice_line_ids)
                    if dc.code in ['61', '112']:
                        amount_general -= totals['amount']
                        exempt_general -= totals['exempt']
                        # tax_general -= doc.amount_tax
                        tax_general -= totals['iva']
                        total_general += abs(doc.amount_total_signed)
                    else:
                        amount_general += totals['amount']
                        exempt_general += totals['exempt']
                        # tax_general += doc.amount_tax
                        tax_general += totals['iva']
                        total_general += abs(doc.amount_total_signed)
                sheet.write(line + 1, 2, abs(amount_general))
                sheet.write(line + 1, 3, abs(exempt_general))
                amount_global += amount_general
                exempt_global += exempt_general
                tax_global += tax_general
            sheet.write(line + 1, 4, abs(tax_general))
            if self.type_operation in ['sell', 'fee']:
                sheet.write(line + 1, 5, total_12_general)
                sheet.write(line + 1, 6, total_general)
            else:
                sheet.write(line + 1, 5, abs(total_general))
            line += 1
        sheet.write(line + 1, 0, "Total General", bold)
        sheet.write(line + 1, 1, total_documents_general, bold)
        sheet.write(line + 1, 2, amount_global, bold)
        sheet.write(line + 1, 3, exempt_global, bold)
        sheet.write(line + 1, 4, tax_global, bold)
        if self.type_operation in ['sell', 'fee']:
            sheet.write(line + 1, 5, total_12_details, bold)
            sheet.write(line + 2, 0, 'Total (Afecto + Excento + Iva + Otros Imp)', bold)
            sheet.write(line + 2, 6, amount_global + exempt_global + tax_global + other_imps, bold)
        else:
            sheet.write(line + 2, 0, 'Total (Afecto + Excento + Iva + Otros Imp)', bold)
            sheet.write(line + 2, 5, amount_global + exempt_global + tax_global + other_imps, bold)

        fp = BytesIO()
        workbook.save(fp)
        fp.seek(0)
        data = fp.read()
        fp.close()
        data_b64 = base64.encodebytes(data)
        attach = self.env['ir.attachment'].create({
            'name': '%s.xls' % (report_name),
            'type': 'binary',
            'datas': data_b64,
        })
        return {
            'type': "ir.actions.act_url",
            'url': "web/content/?model=ir.attachment&id=" + str(
                attach.id) + "&filename_field=name&field=datas&download=true&filename=" + str(attach.name),
            'target': "self",
            'no_destroy': False,
        }

    def print_report(self):
        return self.env.ref('book_account.account_book_sale_action').report_action(self)