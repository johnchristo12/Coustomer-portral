import time
from datetime import datetime

from dateutil.relativedelta import relativedelta
from odoo import fields, models, api, _
from odoo.tools import float_is_zero

import io
import json

try:
    from odoo.tools.misc import xlsxwriter
except ImportError:
    import xlsxwriter


LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

def excel_style(row, col):
    """ Convert given row and column number to an Excel-style cell name. """
    result = []
    while col:
        col, rem = divmod(col-1, 26)
        result[:0] = LETTERS[rem]
    return ''.join(result) + str(row)

class AgeingView(models.TransientModel):
    _inherit = "account.common.report"
    _name = 'account.partner.ageing.soa'

    period_length = fields.Integer(string='Period Length (days)',
                                   required=True, default=30)
    date_from = fields.Date(default=lambda *a: time.strftime('%Y-%m-%d'))
    result_selection = fields.Selection([('customer', 'Receivable Accounts'),
                                         ('supplier', 'Payable Accounts'),
                                         ('customer_supplier',
                                          'Receivable and Payable Accounts')
                                         ], string="Partner's", required=True,
                                        default='customer')

    partner_ids = fields.Many2many(
        'res.partner', string='Partner'
    )
    partner_category_ids = fields.Many2many(
        'res.partner.category', string='Partner Tag',
    )

    @api.model
    def get_partners(self):
        query = """SELECT
                        rp.id as id,
                        rp.name as text
                    FROM
                        res_partner rp
                    WHERE
                        rp.parent_id is null
                        AND rp.active is true"""
        self._cr.execute(query)
        return self._cr.dictfetchall()

    @api.model
    def view_report(self, option):
        r = self.env['account.partner.ageing.soa'].search([('id', '=', option[0])])

        data = {
            'result_selection': r.result_selection,
            'model': self,
            'journals': r.journal_ids,
            'target_move': r.target_move,
            'period_length': r.period_length,
            'partners': r.partner_ids,
            'partner_tags': r.partner_category_ids,

        }
        if r.date_from:
            data.update({
                'date_from': r.date_from,
            })

        filters = self.get_filter(option)

        records = self._get_report_values(data)

        currency = self._get_currency()

        return {
            'name': "Duty Partner Ageing",
            'type': 'ir.actions.client',
            'tag': 'p_a',
            'filters': filters,
            'report_lines': records['Partners'],
            'currency': currency,
        }

    def get_filter(self, option):
        data = self.get_filter_data(option)
        filters = {}

        if data.get('target_move'):
            filters['target_move'] = data.get('target_move')
        if data.get('date_from'):
            filters['date_from'] = data.get('date_from')
        if data.get('result_selection') == 'customer':
            filters['result_selection'] = 'Receivable'
        elif data.get('result_selection') == 'supplier':
            filters['result_selection'] = 'Payable'
        else:
            filters['result_selection'] = 'Receivable and Payable'

        if data.get('partners'):
            filters['partners'] = self.env['res.partner'].browse(
                data.get('partners')).mapped('name')
        else:
            filters['partners'] = ['All']

        if data.get('partner_tags', []):
            filters['partner_tags'] = self.env['res.partner.category'].browse(
                data.get('partner_tags', [])).mapped('name')
        else:
            filters['partner_tags'] = ['All']

        filters['company_id'] = ''
        filters['company_name'] = data.get('company_name')
        filters['partners_list'] = data.get('partners_list')
        filters['category_list'] = data.get('category_list')
        filters['company_name'] = data.get('company_name')
        filters['target_move'] = data.get('target_move').capitalize()


        return filters

    def get_filter_data(self, option):
        r = self.env['account.partner.ageing.soa'].search([('id', '=', option[0])])
        default_filters = {}
        company_id = self.env.company
        company_domain = [('company_id', '=', company_id.id)]
        partner = r.partner_ids if r.partner_ids else self.env[
            'res.partner'].search([])
        categories = r.partner_category_ids if r.partner_category_ids \
            else self.env['res.partner.category'].search([])

        filter_dict = {
            'partners': r.partner_ids.ids,
            'partner_tags': r.partner_category_ids.ids,
            'company_id': company_id.id,
            'date_from': r.date_from,

            'target_move': r.target_move,
            'result_selection': r.result_selection,
            'partners_list': [(p.id, p.name) for p in partner],
            'category_list': [(c.id, c.name) for c in categories],
            'company_name': company_id and company_id.name,
        }
        filter_dict.update(default_filters)
        return filter_dict

    def _get_report_values(self, data):
        docs = data['model']
        date_from = data.get('date_from').strftime('%Y-%m-%d')
        if data['result_selection'] == 'customer':
            account_type = ['receivable']
        elif data['result_selection'] == 'supplier':
            account_type = ['payable']
        else:
            account_type = ['payable', 'receivable']
        target_move = data['target_move']
        partners = data.get('partners')
        if data['partner_tags']:
            partners = self.env['res.partner'].search(
                [('category_id', 'in', data['partner_tags'].ids)])

        account_res = self._get_partner_move_lines(data, partners, date_from,
                                                   target_move,
                                                   account_type,
                                                   data['period_length'])

        return {
            'doc_ids': self.ids,
            'docs': docs,
            'time': time,
            'Partners': account_res,

        }

    @api.model
    def create(self, vals):
        vals['target_move'] = 'posted'
        res = super(AgeingView, self).create(vals)
        return res

    def write(self, vals):
        if vals.get('target_move'):
            vals.update({'target_move': vals.get('target_move').lower()})

        if vals.get('partner_ids'):
            vals.update(
                {'partner_ids': [(4, j) for j in vals.get('partner_ids')]})
        if not vals.get('partner_ids'):
            vals.update({'partner_ids': [(5,)]})
        if vals.get('partner_category_ids'):
            vals.update({'partner_category_ids': [(4, j) for j in vals.get(
                'partner_category_ids')]})
        if not vals.get('partner_category_ids'):
            vals.update({'partner_category_ids': [(5,)]})

        res = super(AgeingView, self).write(vals)
        return res

    def get_duty_vat_details(self, duty_product_ids, vat_product_ids, move_line):
        duty_amount = 0.00
        vat_amount = 0.00
        for line in move_line.move_id.invoice_line_ids:
            if line.product_id and line.product_id.id in duty_product_ids:
                duty_amount += line.price_total
            elif line.product_id and line.product_id.id in vat_product_ids:
                vat_amount += line.price_total
        return duty_amount, vat_amount

    def _get_partner_move_lines(self, data, partners, date_from, target_move,
                                account_type,
                                period_length):

        periods = {}
        duty_product_ids = self.env.company.duty_product_ids.ids
        vat_product_ids = self.env.company.vat_product_ids.ids
        start = datetime.strptime(date_from, "%Y-%m-%d")
        date_from = datetime.strptime(date_from, "%Y-%m-%d").date()
        for i in range(7)[::-1]:
            if i in [6, 5, 4, 3]:
                stop = start - relativedelta(days=period_length)
                period_name = str((7 - (i + 1)) * period_length + 1) + '-' + str(
                    (7 - i) * period_length)
                period_stop = (start - relativedelta(days=1)).strftime('%Y-%m-%d')
            elif i == 2:
                stop = start - relativedelta(days=60)
                period_name = '121-180'
                period_stop = (start - relativedelta(days=1)).strftime('%Y-%m-%d')
            elif i == 1:
                stop = start - relativedelta(days=185)
                period_name = '181-365'
                period_stop = (start - relativedelta(days=1)).strftime('%Y-%m-%d')
            elif i in [0]:
                stop = datetime.strptime('2000-01-01', '%Y-%m-%d')
                period_stop = (start - relativedelta(days=1)).strftime('%Y-%m-%d')
                period_name = '+365'
            periods[str(i)] = {
                'name': period_name,
                'stop': period_stop,
                'start': (i != 0 and stop.strftime('%Y-%m-%d') or False),
            }
            start = stop
        res = []
        total = []
        cr = self.env.cr
        user_company = self.env.company

        user_currency = user_company.currency_id
        ResCurrency = self.env['res.currency'].with_context(date=date_from)
        company_ids = self._context.get('company_ids') or [user_company.id]
        move_state = ['draft', 'posted']
        if target_move == 'posted':
            move_state = ['posted']
        arg_list = (tuple(move_state), tuple(account_type))

        reconciliation_clause = '(l.reconciled IS FALSE)'
        cr.execute(
            'SELECT debit_move_id, credit_move_id FROM account_partial_reconcile where max_date > %s',
            (date_from,))
        reconciled_after_date = []
        for row in cr.fetchall():
            reconciled_after_date += [row[0], row[1]]
        if reconciled_after_date:
            reconciliation_clause = '(l.reconciled IS FALSE OR l.id IN %s)'
            arg_list += (tuple(reconciled_after_date),)

        arg_list += (date_from, tuple(company_ids),)
        partner_list = '(l.partner_id IS NOT  NULL)'
        if partners:
            list = tuple(partners.ids) + tuple([0])
            if list:
                partner_list = '(l.partner_id IS NULL OR l.partner_id IN %s)'
                arg_list += (tuple(list),)
        query = '''
                    SELECT DISTINCT l.partner_id, UPPER(res_partner.name)
                    FROM account_move_line AS l left join res_partner on l.partner_id = res_partner.id, account_account, account_move am
                    WHERE (l.account_id = account_account.id)
                        AND (l.move_id = am.id)
                        AND (am.state IN %s)
                        AND (account_account.internal_type IN %s)
                       
                        AND ''' + reconciliation_clause + '''          
                        AND (l.date <= %s)
                        AND l.company_id IN %s
                        AND ''' + partner_list + '''
                           
                    ORDER BY UPPER(res_partner.name)'''
        cr.execute(query, arg_list)


        partners = cr.dictfetchall()

        # put a total of 0
        for i in range(9):
            total.append(0)

        # Build a string like (1,2,3) for easy use in SQL query
        partner_ids = [partner['partner_id'] for partner in partners if
                       partner['partner_id']]

        lines = dict(
            (partner['partner_id'] or False, []) for partner in partners)
        if not partner_ids:
            return [], [], {}

        # This dictionary will store the not due amount of all partners
        undue_amounts = {}
        undue_paid_amount = {}
        query = '''SELECT l.id
                        FROM account_move_line AS l, account_account, account_move am
                        WHERE (l.account_id = account_account.id) AND (l.move_id = am.id)
                            AND (am.state IN %s)
                            AND (account_account.internal_type IN %s)
                            AND (l.date >= %s)\
                            AND ((l.partner_id IN %s) OR (l.partner_id IS NULL))
                        AND (l.date <= %s)
                        AND l.company_id IN %s'''
        cr.execute(query, (
            tuple(move_state), tuple(account_type), date_from,
            tuple(partner_ids), date_from, tuple(company_ids)))
        aml_ids = cr.fetchall()
        aml_ids = aml_ids and [x[0] for x in aml_ids] or []
        for line in self.env['account.move.line'].browse(aml_ids):
            partner_id = line.partner_id.id or False
            move_id = line.move_id.id
            move_name = line.move_id.name
            date_maturity = line.date
            account_id = line.account_id.name
            account_code = line.account_id.code
            jrnl_id = line.journal_id.name
            currency_id = line.company_id.currency_id.position
            currency_symbol = line.company_id.currency_id.symbol

            if partner_id not in undue_amounts:
                undue_amounts[partner_id] = {
                    'amount': 0.00,
                    'duty_amount': 0.00,
                    'vat_amount': 0.00}
            if partner_id not in undue_paid_amount:
                undue_paid_amount[partner_id] = 0.00
            line_amount = ResCurrency._compute(line.company_id.currency_id,
                                               user_currency, line.balance)
            if user_currency.is_zero(line_amount):
                continue
            for partial_line in line.matched_debit_ids:
                if partial_line.max_date <= date_from and partial_line.debit_move_id.move_id.state == 'posted' and \
                            partial_line.credit_move_id.move_id.state == 'posted':
                    line_amount += ResCurrency._compute(
                        partial_line.company_id.currency_id, user_currency,
                        partial_line.amount)
            for partial_line in line.matched_credit_ids:
                if partial_line.max_date <= date_from and partial_line.debit_move_id.move_id.state == 'posted' and \
                            partial_line.credit_move_id.move_id.state == 'posted':
                    line_amount -= ResCurrency._compute(
                        partial_line.company_id.currency_id, user_currency,
                        partial_line.amount)
            if not self.env.company.currency_id.is_zero(line_amount):
                paid_amount = 0.00
                invoice_amount = 0.00
                duty_amount, vat_amount = self.get_duty_vat_details(duty_product_ids, vat_product_ids, line)
                if line.account_id.user_type_id.type == 'receivable':
                    if line_amount < 0.00:
                        paid_amount += line_amount
                        undue_paid_amount[partner_id] += line_amount
                    else:
                        invoice_amount += line_amount
                        undue_amounts[partner_id]['amount'] += line_amount
                        undue_amounts[partner_id]['duty_amount'] += duty_amount
                        undue_amounts[partner_id]['vat_amount'] += vat_amount
                elif line.account_id.user_type_id.type == 'payable':
                    if line_amount > 0.00:
                        paid_amount += line_amount
                        undue_paid_amount[partner_id] += line_amount
                    else:
                        invoice_amount += line_amount
                        undue_amounts[partner_id]['amount'] += line_amount
                        undue_amounts[partner_id]['duty_amount'] += duty_amount
                        undue_amounts[partner_id]['vat_amount'] += vat_amount
                lines[partner_id].append({
                    'line': line,
                    'partner_id': partner_id,
                    'move': move_name,
                    'jrnl': jrnl_id,
                    'currency': currency_id,
                    'symbol': currency_symbol,
                    'acc_name': account_id,
                    'mov_id': move_id,
                    'acc_code': account_code,
                    'date': date_maturity,
                    'amount': invoice_amount,
                    'paid_amount': paid_amount,
                    'duty_amount': duty_amount,
                    'vat_amount': vat_amount,
                    'period7': 7,
                })

        # Use one query per period and store results in history (a list variable)
        # Each history will contain: history[1] = {'<partner_id>': <partner_debit-credit>}
        history = []
        duty_hist = []
        vat_hist = []
        for i in range(7):
            args_list = (
                tuple(move_state), tuple(account_type), tuple(partner_ids),)
            dates_query = '(l.date'

            if periods[str(i)]['start'] and periods[str(i)]['stop']:
                dates_query += ' BETWEEN %s AND %s)'

                args_list += (
                    periods[str(i)]['start'], periods[str(i)]['stop'])
            elif periods[str(i)]['start']:
                dates_query += ' >= %s)'

                args_list += (periods[str(i)]['start'],)
            else:
                dates_query += ' <= %s)'
                args_list += (periods[str(i)]['stop'],)

            args_list += (date_from, tuple(company_ids))

            query = '''SELECT l.id
                            FROM account_move_line AS l, account_account, account_move am
                            WHERE (l.account_id = account_account.id) AND (l.move_id = am.id)
                                AND (am.state IN %s)
                                AND (account_account.internal_type IN %s)
                                AND ((l.partner_id IN %s) OR (l.partner_id IS NULL))
                                AND ''' + dates_query + '''
                                
                                
                            AND (l.date <= %s)
                            AND l.company_id IN %s'''
            cr.execute(query, args_list)

            partners_amount = {}
            duty_amount_dict = {}
            vat_amount_dict = {}
            aml_ids = cr.fetchall()
            aml_ids = aml_ids and [x[0] for x in aml_ids] or []
            for line in self.env['account.move.line'].browse(aml_ids):
                partner_id = line.partner_id.id or False
                move_id = line.move_id.id
                move_name = line.move_id.name
                date_maturity = line.date.strftime("%d/%m/%Y")
                account_id = line.account_id.name
                account_code = line.account_id.code
                jrnl_id = line.journal_id.name
                currency_id = line.company_id.currency_id.position
                currency_symbol = line.company_id.currency_id.symbol
                if partner_id not in partners_amount:
                    partners_amount[partner_id] = 0.0
                    duty_amount_dict[partner_id] = 0.00
                    vat_amount_dict[partner_id] = 0.00
                if partner_id not in undue_paid_amount:
                    undue_paid_amount[partner_id] = 0.00
                line_amount = ResCurrency._compute(line.company_id.currency_id,
                                                   user_currency, line.balance)
                if user_currency.is_zero(line_amount):
                    continue
                for partial_line in line.matched_debit_ids:
                    if partial_line.max_date <= date_from and partial_line.debit_move_id.move_id.state == 'posted' and \
                            partial_line.credit_move_id.move_id.state == 'posted':
                        line_amount += ResCurrency._compute(
                            partial_line.company_id.currency_id, user_currency,
                            partial_line.amount)
                for partial_line in line.matched_credit_ids:
                    if partial_line.max_date <= date_from and partial_line.debit_move_id.move_id.state == 'posted' and \
                            partial_line.credit_move_id.move_id.state == 'posted':
                        line_amount -= ResCurrency._compute(
                            partial_line.company_id.currency_id, user_currency,
                            partial_line.amount)

                if not self.env.company.currency_id.is_zero(
                        line_amount):
                    invoice_amount = 0.00
                    paid_amount = 0.00
                    if line.account_id.user_type_id.type == 'receivable':
                        if line_amount < 0.00:
                            paid_amount += line_amount
                        else:
                            invoice_amount += line_amount
                    elif line.account_id.user_type_id.type == 'payable':
                        if line_amount > 0.00:
                            paid_amount += line_amount
                        else:
                            invoice_amount += line_amount
                    partners_amount[partner_id] += invoice_amount
                    undue_paid_amount[partner_id] += paid_amount
                    duty_amount, vat_amount = self.get_duty_vat_details(duty_product_ids, vat_product_ids, line)
                    duty_amount_dict[partner_id] += duty_amount
                    vat_amount_dict[partner_id] += vat_amount
                    if i + 1 == 7:
                        period7 = i + 1
                        lines[partner_id].append({
                            'period7': period7,
                            'line': line,
                            'partner_id': partner_id,
                            'move': move_name,
                            'currency': currency_id,
                            'symbol': currency_symbol,
                            'jrnl': jrnl_id,
                            'acc_name': account_id,
                            'mov_id': move_id,
                            'acc_code': account_code,
                            'date': date_maturity,
                            'amount': invoice_amount,
                            'paid_amount': paid_amount,
                            'duty_amount': duty_amount,
                            'vat_amount': vat_amount
                        })
                    elif i + 1 == 6:
                        period6 = i + 1
                        lines[partner_id].append({
                            'period6': period6,
                            'line': line,
                            'partner_id': partner_id,
                            'move': move_name,
                            'currency': currency_id,
                            'symbol': currency_symbol,
                            'jrnl': jrnl_id,
                            'acc_name': account_id,
                            'mov_id': move_id,
                            'acc_code': account_code,
                            'date': date_maturity,
                            'amount': invoice_amount,
                            'paid_amount': paid_amount,
                            'duty_amount': duty_amount,
                            'vat_amount': vat_amount
                        })
                    elif i + 1 == 5:
                        period5 = i + 1
                        lines[partner_id].append({
                            'period5': period5,
                            'line': line,
                            'partner_id': partner_id,
                            'move': move_name,
                            'currency': currency_id,
                            'symbol': currency_symbol,
                            'jrnl': jrnl_id,
                            'acc_name': account_id,
                            'mov_id': move_id,
                            'acc_code': account_code,
                            'date': date_maturity,
                            'amount': invoice_amount,
                            'paid_amount': paid_amount,
                            'duty_amount': duty_amount,
                            'vat_amount': vat_amount
                        })
                    elif i + 1 == 4:
                        period4 = i + 1
                        lines[partner_id].append({

                            'period4': period4,
                            'line': line,
                            'partner_id': partner_id,
                            'move': move_name,
                            'jrnl': jrnl_id,
                            'acc_name': account_id,
                            'currency': currency_id,
                            'symbol': currency_symbol,
                            'mov_id': move_id,
                            'acc_code': account_code,
                            'date': date_maturity,
                            'amount': invoice_amount,
                            'paid_amount': paid_amount,
                            'duty_amount': duty_amount,
                            'vat_amount': vat_amount
                        })
                    elif i + 1 == 3:
                        period3 = i + 1
                        lines[partner_id].append({

                            'period3': period3,
                            'line': line,
                            'partner_id': partner_id,
                            'move': move_name,
                            'jrnl': jrnl_id,
                            'acc_name': account_id,
                            'currency': currency_id,
                            'symbol': currency_symbol,
                            'mov_id': move_id,
                            'acc_code': account_code,
                            'date': date_maturity,
                            'amount': invoice_amount,
                            'paid_amount': paid_amount,
                            'duty_amount': duty_amount,
                            'vat_amount': vat_amount
                        })
                    elif i + 1 == 2:
                        period2 = i + 1
                        lines[partner_id].append({

                            'period2': period2,
                            'line': line,
                            'partner_id': partner_id,
                            'move': move_name,
                            'jrnl': jrnl_id,
                            'acc_name': account_id,
                            'currency': currency_id,
                            'symbol': currency_symbol,
                            'mov_id': move_id,
                            'acc_code': account_code,
                            'date': date_maturity,
                            'amount': invoice_amount,
                            'paid_amount': paid_amount,
                            'duty_amount': duty_amount,
                            'vat_amount': vat_amount
                        })
                    else:
                        period1 = i + 1
                        lines[partner_id].append({

                            'period1': period1,
                            'line': line,
                            'partner_id': partner_id,
                            'move': move_name,
                            'jrnl': jrnl_id,
                            'acc_name': account_id,
                            'currency': currency_id,
                            'symbol': currency_symbol,
                            'mov_id': move_id,
                            'acc_code': account_code,
                            'date': date_maturity,
                            'amount': invoice_amount,
                            'paid_amount': paid_amount,
                            'duty_amount': duty_amount,
                            'vat_amount': vat_amount
                        })

            history.append(partners_amount)
            duty_hist.append(duty_amount_dict)
            vat_hist.append(vat_amount_dict)

        for partner in partners:
            if partner['partner_id'] is None:
                partner['partner_id'] = False
            at_least_one_amount = False
            values = {}
            undue_amt = 0.0
            undue_duty_amount = 0.00
            undue_vat_amount = 0.00
            paid_amount = 0.00
            if partner['partner_id'] in undue_amounts:  # Making sure this partner actually was found by the query
                undue_amt = undue_amounts[partner['partner_id']]['amount']
                undue_duty_amount = undue_amounts[partner['partner_id']]['duty_amount']
                undue_vat_amount = undue_amounts[partner['partner_id']]['vat_amount']
            if partner['partner_id'] in undue_paid_amount:
                paid_amount = undue_paid_amount[partner['partner_id']]
                if paid_amount != 0.00:
                    at_least_one_amount = True
            total[8] = total[8] + undue_amt + paid_amount
            values['direction'] = 0.00
            values['unalloc'] = paid_amount
            for rec in lines:
                if partner['partner_id'] == rec:
                    child_lines = lines[rec]
            values['child_lines'] = child_lines
            if not float_is_zero(values['direction'],
                                 precision_rounding=self.env.company.currency_id.rounding):
                at_least_one_amount = True

            for i in range(7):
                during = False
                if partner['partner_id'] in history[i]:
                    during = [history[i][partner['partner_id']]]
                    duty_during = [duty_hist[i][partner['partner_id']]]
                    vat_during = [vat_hist[i][partner['partner_id']]]
                # Adding counter
                total[(i)] = total[(i)] + (during and during[0] or 0)
                values[str(i)] = {
                    'amount': during and during[0] or 0.0,
                    'duty_amount': duty_during and duty_during[0] or 0.00,
                    'vat_amount': vat_during and vat_during[0] or 0.00
                    }
                if not float_is_zero(values[str(i)]['amount'],
                                     precision_rounding=self.env.company.currency_id.rounding):
                    at_least_one_amount = True
            values['6']['amount'] += undue_amt
            values['6']['duty_amount'] += undue_duty_amount
            values['6']['vat_amount'] += undue_vat_amount
            values['total'] = sum(
                [values['unalloc']] + [values['direction']] + [values[str(i)]['amount'] for i in range(7)])
            ## Add for total
            total[(i + 1)] += values['total']
            values['partner_id'] = partner['partner_id']
            if partner['partner_id']:
                browsed_partner = self.env['res.partner'].browse(
                    partner['partner_id'])
                values['name'] = browsed_partner.name and len(
                    browsed_partner.name) >= 45 and browsed_partner.name[
                                                    0:40] + '...' or browsed_partner.name
                values['trust'] = browsed_partner.trust
            else:
                values['name'] = _('Unknown Partner')
                values['trust'] = False

            if at_least_one_amount or (
                    self._context.get('include_nullified_amount') and lines[
                partner['partner_id']]):
                res.append(values)
        return res, total, lines

    @api.model
    def _get_currency(self):
        journal = self.env['account.journal'].browse(
            self.env.context.get('default_journal_id', False))
        if journal.currency_id:
            return journal.currency_id.id
        lang = self.env.user.lang
        if not lang:
            lang = 'en_US'
        lang = lang.replace("_", '-')
        currency_array = [self.env.company.currency_id.symbol,
                          self.env.company.currency_id.position,
                          lang,
                          self.env.company.currency_id.decimal_places]
        return currency_array

    def get_dynamic_xlsx_report(self, data, response, report_data, dfr_data ):

        report_data_main = json.loads(report_data)
        output = io.BytesIO()

        filters = json.loads(data)

        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Partner Aging - Summary')
         ##FORMATS##
        heading_format = workbook.add_format({'align': 'center',
                                              'valign': 'vcenter',
                                              'bold': True, 'size': 18})
        sub_heading_format = workbook.add_format({'align': 'center',
                                                  'valign': 'vcenter',
                                                  'bold': True, 'size': 14})
        bold = workbook.add_format({'bold': True})
        bold_center = workbook.add_format({'bold': True, 'valign': 'vcenter', 'bg_color': '#b5b5b5'})
        bold_center_center = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#b5b5b5'})
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
        no_format = workbook.add_format({'num_format': '#,##0.00'})
        normal_num_bold = workbook.add_format({'bold': True, 'num_format': '#,##0.00'})
        ##FORMATS ENDS##
        worksheet.merge_range('A1:J1', "Partner Aging - Summary", sub_heading_format)
        row = 1
        if filters.get('date_from'):
            worksheet.write(row, 0, 'Date Ason', bold)
            worksheet.write_datetime(row, 1, datetime.strptime(filters.get('date_from'), "%Y-%m-%d"), date_format)
            row += 1
        worksheet.write(row, 0, 'Account Type', bold)
        filter_from = excel_style(row + 1, 2)
        filter_to = excel_style(row + 1, 10)
        worksheet.merge_range('%s:%s'%(filter_from, filter_to), filters.get('result_selection'), bold)
        row += 1
        worksheet.write(row, 0, 'Target Moves', bold)
        filter_from = excel_style(row + 1, 2)
        filter_to = excel_style(row + 1, 10)
        worksheet.merge_range('%s:%s'%(filter_from, filter_to), filters.get('target_move'), bold)
        row += 1
        worksheet.write(row, 0, 'Partners', bold)
        filter_from = excel_style(row + 1, 2)
        filter_to = excel_style(row + 1, 10)
        worksheet.merge_range('%s:%s'%(filter_from, filter_to), ', '.join([lt or '' for lt in filters['partners']]), bold)
        row += 1
        worksheet.write(row, 0, 'Partner Type', bold)
        filter_from = excel_style(row + 1, 2)
        filter_to = excel_style(row + 1, 10)
        worksheet.merge_range('%s:%s'%(filter_from, filter_to), ', '.join([lt or '' for lt in filters['partner_tags']]), bold)
        row += 2
        
        worksheet.set_column('A:A', 35)
        worksheet.set_column('B:B', 25)
        worksheet.set_column('C:C', 50)
        worksheet.set_column('D:D', 25)
        worksheet.set_column('E:E', 25)
        worksheet.set_column('F:F', 25)
        worksheet.set_column('G:G', 25)
        worksheet.set_column('H:H', 25)
        worksheet.set_column('I:I', 25)
        worksheet.set_column('J:J', 25)
        worksheet.set_column('K:K', 25)
        worksheet.set_column('L:L', 25)
        worksheet.set_column('M:M', 25)
        worksheet.set_column('N:N', 25)
        worksheet.set_column('O:O', 25)
        worksheet.set_column('P:P', 25)
        worksheet.set_column('Q:Q', 25)
        worksheet.set_column('R:R', 25)
        worksheet.set_column('S:S', 25)
        worksheet.set_column('T:T', 25)
        worksheet.set_column('U:U', 25)
        worksheet.set_column('V:V', 25)
        worksheet.set_column('W:W', 25)
        worksheet.set_column('X:X', 25)
        worksheet.set_column('Y:Y', 25)
        worksheet.set_column('Z:Z', 25)
        worksheet.write(row, 0, 'Partner', bold_center)
        worksheet.write(row, 1, 'Unallocated', bold_center)
        mg_from = excel_style(row + 1, 3)
        mg_to = excel_style(row + 1, 5)
        worksheet.merge_range('%s:%s'%(mg_from, mg_to), '0-30', bold_center)
        mg_from = excel_style(row + 1, 6)
        mg_to = excel_style(row + 1, 8)
        worksheet.merge_range('%s:%s'%(mg_from, mg_to), '31-60', bold_center)
        mg_from = excel_style(row + 1, 9)
        mg_to = excel_style(row + 1, 11)
        worksheet.merge_range('%s:%s'%(mg_from, mg_to), '61-90', bold_center)
        mg_from = excel_style(row + 1, 12)
        mg_to = excel_style(row + 1, 14)
        worksheet.merge_range('%s:%s'%(mg_from, mg_to), '91-120', bold_center)
        mg_from = excel_style(row + 1, 15)
        mg_to = excel_style(row + 1, 17)
        worksheet.merge_range('%s:%s'%(mg_from, mg_to), '121-180', bold_center)
        mg_from = excel_style(row + 1, 18)
        mg_to = excel_style(row + 1, 20)
        worksheet.merge_range('%s:%s'%(mg_from, mg_to), '181-365', bold_center)
        mg_from = excel_style(row + 1, 21)
        mg_to = excel_style(row + 1, 23)
        worksheet.merge_range('%s:%s'%(mg_from, mg_to), '365+', bold_center)
        mg_from = excel_style(row + 1, 24)
        mg_to = excel_style(row + 1, 26)
        worksheet.merge_range('%s:%s'%(mg_from, mg_to), 'Total', bold_center)
        row += 1
        worksheet.write(row, 2, 'Total', bold_center)
        worksheet.write(row, 3, 'Duty Amount', bold_center)
        worksheet.write(row, 4, 'Tax Amount', bold_center)
        worksheet.write(row, 5, 'Total', bold_center)
        worksheet.write(row, 6, 'Duty Amount', bold_center)
        worksheet.write(row, 7, 'Tax Amount', bold_center)
        worksheet.write(row, 8, 'Total', bold_center)
        worksheet.write(row, 9, 'Duty Amount', bold_center)
        worksheet.write(row, 10, 'Tax Amount', bold_center)
        worksheet.write(row, 11, 'Total', bold_center)
        worksheet.write(row, 12, 'Duty Amount', bold_center)
        worksheet.write(row, 13, 'Tax Amount', bold_center)
        worksheet.write(row, 14, 'Total', bold_center)
        worksheet.write(row, 15, 'Duty Amount', bold_center)
        worksheet.write(row, 16, 'Tax Amount', bold_center)
        worksheet.write(row, 17, 'Total', bold_center)
        worksheet.write(row, 18, 'Duty Amount', bold_center)
        worksheet.write(row, 19, 'Tax Amount', bold_center)
        worksheet.write(row, 20, 'Total', bold_center)
        worksheet.write(row, 21, 'Duty Amount', bold_center)
        worksheet.write(row, 22, 'Tax Amount', bold_center)
        worksheet.write(row, 23, 'Total', bold_center)
        worksheet.write(row, 24, 'Duty Amount', bold_center)
        worksheet.write(row, 25, 'Tax Amount', bold_center)
        row += 1
        unall = 0.00
        due_30 = 0.00
        due_30_duty = 0.00
        due_30_vat = 0.00
        due_60 = 0.00
        due_60_duty = 0.00
        due_60_vat = 0.00
        due_90 = 0.00
        due_90_duty = 0.00
        due_90_vat = 0.00
        due_120 = 0.00
        due_120_duty = 0.00
        due_120_vat = 0.00
        due_180 = 0.00
        due_180_duty = 0.00
        due_180_vat = 0.00
        due_365 = 0.00
        due_365_duty = 0.00
        due_365_vat = 0.00
        due_365_plus = 0.00
        due_365_p_duty = 0.00
        due_365_p_vat = 0.00
        total = 0.00
        total_duty = 0.00
        total_vat = 0.00
        for rec_data in report_data_main[0]:
            line_duty_total = 0.00
            line_vat_total = 0.00
            worksheet.write(row, 0, rec_data['name'])
            worksheet.write_number(row, 1, rec_data['unalloc'], no_format)
            worksheet.write_number(row, 2, rec_data['6']['amount'], no_format)
            worksheet.write_number(row, 3, rec_data['6']['duty_amount'], no_format)
            worksheet.write_number(row, 4, rec_data['6']['vat_amount'], no_format)
            line_duty_total += rec_data['6']['duty_amount']
            line_vat_total += rec_data['6']['vat_amount']
            worksheet.write_number(row, 5, rec_data['5']['amount'], no_format)
            worksheet.write_number(row, 6, rec_data['5']['duty_amount'], no_format)
            worksheet.write_number(row, 7, rec_data['5']['vat_amount'], no_format)
            line_duty_total += rec_data['5']['duty_amount']
            line_vat_total += rec_data['5']['vat_amount']
            worksheet.write_number(row, 8, rec_data['4']['amount'], no_format)
            worksheet.write_number(row, 9, rec_data['4']['duty_amount'], no_format)
            worksheet.write_number(row, 10, rec_data['4']['vat_amount'], no_format)
            line_duty_total += rec_data['4']['duty_amount']
            line_vat_total += rec_data['4']['vat_amount']
            worksheet.write_number(row, 11, rec_data['3']['amount'], no_format)
            worksheet.write_number(row, 12, rec_data['3']['duty_amount'], no_format)
            worksheet.write_number(row, 13, rec_data['3']['vat_amount'], no_format)
            line_duty_total += rec_data['3']['duty_amount']
            line_vat_total += rec_data['3']['vat_amount']
            worksheet.write_number(row, 14, rec_data['2']['amount'], no_format)
            worksheet.write_number(row, 15, rec_data['2']['duty_amount'], no_format)
            worksheet.write_number(row, 16, rec_data['2']['vat_amount'], no_format)
            line_duty_total += rec_data['2']['duty_amount']
            line_vat_total += rec_data['2']['vat_amount']
            worksheet.write_number(row, 17, rec_data['1']['amount'], no_format)
            worksheet.write_number(row, 18, rec_data['1']['duty_amount'], no_format)
            worksheet.write_number(row, 19, rec_data['1']['vat_amount'], no_format)
            line_duty_total += rec_data['1']['duty_amount']
            line_vat_total += rec_data['1']['vat_amount']
            worksheet.write_number(row, 20, rec_data['0']['amount'], no_format)
            worksheet.write_number(row, 21, rec_data['0']['duty_amount'], no_format)
            worksheet.write_number(row, 22, rec_data['0']['vat_amount'], no_format)
            line_duty_total += rec_data['0']['duty_amount']
            line_vat_total += rec_data['0']['vat_amount']
            worksheet.write_number(row, 23, rec_data['total'], no_format)
            worksheet.write_number(row, 24, line_duty_total, no_format)
            worksheet.write_number(row, 25, line_vat_total, no_format)
            unall += rec_data['unalloc']
            due_30 += rec_data['6']['amount']
            due_30_duty += rec_data['6']['duty_amount']
            due_30_vat += rec_data['6']['vat_amount']
            due_60 += rec_data['5']['amount']
            due_60_duty += rec_data['5']['duty_amount']
            due_60_vat += rec_data['5']['vat_amount']
            due_90 += rec_data['4']['amount']
            due_90_duty += rec_data['4']['duty_amount']
            due_90_vat += rec_data['4']['vat_amount']
            due_120 += rec_data['3']['amount']
            due_120_duty += rec_data['3']['duty_amount']
            due_120_vat += rec_data['3']['vat_amount']
            due_180 += rec_data['2']['amount']
            due_180_duty += rec_data['2']['duty_amount']
            due_180_vat += rec_data['2']['vat_amount']
            due_365 += rec_data['1']['amount']
            due_365_duty += rec_data['1']['duty_amount']
            due_365_vat += rec_data['1']['vat_amount']
            due_365_plus += rec_data['0']['amount']
            due_365_p_duty += rec_data['0']['duty_amount']
            due_365_p_vat += rec_data['0']['vat_amount']
            total += rec_data['total']
            total_duty += line_duty_total
            total_vat += line_vat_total
            row += 1
        worksheet.write(row, 0, "Total", bold)
        worksheet.write_number(row, 1, unall, normal_num_bold)
        worksheet.write_number(row, 2, due_30, normal_num_bold)
        worksheet.write_number(row, 3, due_30_duty, normal_num_bold)
        worksheet.write_number(row, 4, due_30_vat, normal_num_bold)
        worksheet.write_number(row, 5, due_60, normal_num_bold)
        worksheet.write_number(row, 6, due_60_duty, normal_num_bold)
        worksheet.write_number(row, 7, due_60_vat, normal_num_bold)
        worksheet.write_number(row, 8, due_90, normal_num_bold)
        worksheet.write_number(row, 9, due_90_duty, normal_num_bold)
        worksheet.write_number(row, 10, due_90_vat, normal_num_bold)
        worksheet.write_number(row, 11, due_120, normal_num_bold)
        worksheet.write_number(row, 12, due_120_duty, normal_num_bold)
        worksheet.write_number(row, 13, due_120_vat, normal_num_bold)
        worksheet.write_number(row, 14, due_180, normal_num_bold)
        worksheet.write_number(row, 15, due_180_duty, normal_num_bold)
        worksheet.write_number(row, 16, due_180_vat, normal_num_bold)
        worksheet.write_number(row, 17, due_365, normal_num_bold)
        worksheet.write_number(row, 18, due_365_duty, normal_num_bold)
        worksheet.write_number(row, 19, due_365_vat, normal_num_bold)
        worksheet.write_number(row, 20, due_365_plus, normal_num_bold)
        worksheet.write_number(row, 21, due_365_p_duty, normal_num_bold)
        worksheet.write_number(row, 22, due_365_p_vat, normal_num_bold)
        worksheet.write_number(row, 23, total, normal_num_bold)
        worksheet.write_number(row, 24, total_duty, normal_num_bold)
        worksheet.write_number(row, 25, total_vat, normal_num_bold)
        row += 5
        summ_from = excel_style(row + 1, 1)
        summ_to = excel_style(row + 1, 3)
        worksheet.merge_range('%s:%s'%(summ_from, summ_to), "Summary", bold_center_center)
        row += 1
        worksheet.write(row, 0, 'Period', bold_center)
        worksheet.write(row, 1, 'Amount', bold_center)
        worksheet.write(row, 2, '%', bold_center)
        row += 1
        worksheet.write(row, 0, '0-30')
        worksheet.write_number(row, 1, due_30, no_format)
        worksheet.write_number(row, 2, round(((due_30 / (total - unall)) * 100.00), 2), no_format)
        row += 1
        worksheet.write(row, 0, '31-60')
        worksheet.write_number(row, 1, due_60, no_format)
        worksheet.write_number(row, 2, round(((due_60 / (total - unall)) * 100.00), 2), no_format)
        row += 1
        worksheet.write(row, 0, '61-90')
        worksheet.write_number(row, 1, due_90, no_format)
        worksheet.write_number(row, 2, round(((due_90 / (total - unall)) * 100.00), 2), no_format)
        row += 1
        worksheet.write(row, 0, '91-120')
        worksheet.write_number(row, 1, due_120, no_format)
        worksheet.write_number(row, 2, round(((due_120 / (total - unall)) * 100.00), 2), no_format)
        row += 1
        worksheet.write(row, 0, '121-180')
        worksheet.write_number(row, 1, due_180, no_format)
        worksheet.write_number(row, 2, round(((due_180 / (total - unall)) * 100.00), 2), no_format)
        row += 1
        worksheet.write(row, 0, '181-365')
        worksheet.write_number(row, 1, due_365, no_format)
        worksheet.write_number(row, 2, round(((due_365 / (total - unall)) * 100.00), 2), no_format)
        row += 1
        worksheet.write(row, 0, '365 +')
        worksheet.write_number(row, 1, due_365_plus, no_format)
        worksheet.write_number(row, 2, round(((due_365_plus / (total - unall)) * 100.00), 2), no_format)
        row += 1
        worksheet.write(row, 0, 'Total :', bold)
        worksheet.write_number(row, 1, total, normal_num_bold)
        row += 1
        worksheet.write(row, 0, 'Unallocated')
        worksheet.write_number(row, 1, unall, no_format)
        worksheet = workbook.add_worksheet('Partner Aging - Detail')
        worksheet.merge_range('A1:M1', "Partner Aging - Detail", sub_heading_format)
        row = 1
        if filters.get('date_from'):
            worksheet.write(row, 0, 'Date Ason', bold)
            worksheet.write_datetime(row, 1, datetime.strptime(filters.get('date_from'), "%Y-%m-%d"), date_format)
            row += 1
        worksheet.write(row, 0, 'Account Type', bold)
        filter_from = excel_style(row + 1, 2)
        filter_to = excel_style(row + 1, 10)
        worksheet.merge_range('%s:%s'%(filter_from, filter_to), filters.get('result_selection'), bold)
        row += 1
        worksheet.write(row, 0, 'Target Moves', bold)
        filter_from = excel_style(row + 1, 2)
        filter_to = excel_style(row + 1, 10)
        worksheet.merge_range('%s:%s'%(filter_from, filter_to), filters.get('target_move'), bold)
        row += 1
        worksheet.write(row, 0, 'Partners', bold)
        filter_from = excel_style(row + 1, 2)
        filter_to = excel_style(row + 1, 10)
        worksheet.merge_range('%s:%s'%(filter_from, filter_to), ', '.join([lt or '' for lt in filters['partners']]), bold)
        row += 1
        worksheet.write(row, 0, 'Partner Type', bold)
        filter_from = excel_style(row + 1, 2)
        filter_to = excel_style(row + 1, 10)
        worksheet.merge_range('%s:%s'%(filter_from, filter_to), ', '.join([lt or '' for lt in filters['partner_tags']]), bold)
        row += 2
        
        worksheet.set_column('A:A', 35)
        worksheet.set_column('B:B', 25)
        worksheet.set_column('C:C', 50)
        worksheet.set_column('D:D', 25)
        worksheet.set_column('E:E', 25)
        worksheet.set_column('F:F', 25)
        worksheet.set_column('G:G', 25)
        worksheet.set_column('H:H', 25)
        worksheet.set_column('I:I', 25)
        worksheet.set_column('J:J', 25)
        worksheet.set_column('K:K', 25)
        worksheet.set_column('L:L', 25)
        worksheet.set_column('M:M', 25)
        worksheet.set_column('N:N', 25)
        worksheet.set_column('O:O', 25)
        worksheet.set_column('P:P', 25)
        worksheet.set_column('Q:Q', 25)
        worksheet.set_column('R:R', 25)
        worksheet.set_column('S:S', 25)
        worksheet.set_column('T:T', 25)
        worksheet.set_column('U:U', 25)
        worksheet.set_column('V:V', 25)
        worksheet.set_column('W:W', 25)
        worksheet.set_column('X:X', 25)
        worksheet.set_column('Y:Y', 25)
        worksheet.set_column('Z:Z', 25)
        worksheet.write(row, 0, 'Entry Label', bold_center)
        worksheet.write(row, 1, 'Due Date', bold_center)
        worksheet.write(row, 2, 'Journal', bold_center)
        worksheet.write(row, 3, 'Account', bold_center)
        worksheet.write(row, 4, 'Unallocated', bold_center)
        mg_from = excel_style(row + 1, 6)
        mg_to = excel_style(row + 1, 8)
        worksheet.merge_range('%s:%s'%(mg_from, mg_to), '0-30', bold_center)
        mg_from = excel_style(row + 1, 9)
        mg_to = excel_style(row + 1, 11)
        worksheet.merge_range('%s:%s'%(mg_from, mg_to), '31-60', bold_center)
        mg_from = excel_style(row + 1, 12)
        mg_to = excel_style(row + 1, 14)
        worksheet.merge_range('%s:%s'%(mg_from, mg_to), '61-90', bold_center)
        mg_from = excel_style(row + 1, 15)
        mg_to = excel_style(row + 1, 17)
        worksheet.merge_range('%s:%s'%(mg_from, mg_to), '91-120', bold_center)
        mg_from = excel_style(row + 1, 18)
        mg_to = excel_style(row + 1, 20)
        worksheet.merge_range('%s:%s'%(mg_from, mg_to), '121-180', bold_center)
        mg_from = excel_style(row + 1, 21)
        mg_to = excel_style(row + 1, 23)
        worksheet.merge_range('%s:%s'%(mg_from, mg_to), '181-365', bold_center)
        mg_from = excel_style(row + 1, 24)
        mg_to = excel_style(row + 1, 26)
        worksheet.merge_range('%s:%s'%(mg_from, mg_to), '365+', bold_center)
        row += 1
        worksheet.write(row, 5, 'Total', bold_center)
        worksheet.write(row, 6, 'Duty Amount', bold_center)
        worksheet.write(row, 7, 'Tax Amount', bold_center)
        worksheet.write(row, 8, 'Total', bold_center)
        worksheet.write(row, 9, 'Duty Amount', bold_center)
        worksheet.write(row, 10, 'Tax Amount', bold_center)
        worksheet.write(row, 11, 'Total', bold_center)
        worksheet.write(row, 12, 'Duty Amount', bold_center)
        worksheet.write(row, 13, 'Tax Amount', bold_center)
        worksheet.write(row, 14, 'Total', bold_center)
        worksheet.write(row, 15, 'Duty Amount', bold_center)
        worksheet.write(row, 16, 'Tax Amount', bold_center)
        worksheet.write(row, 17, 'Total', bold_center)
        worksheet.write(row, 18, 'Duty Amount', bold_center)
        worksheet.write(row, 19, 'Tax Amount', bold_center)
        worksheet.write(row, 20, 'Total', bold_center)
        worksheet.write(row, 21, 'Duty Amount', bold_center)
        worksheet.write(row, 22, 'Tax Amount', bold_center)
        worksheet.write(row, 23, 'Total', bold_center)
        worksheet.write(row, 24, 'Duty Amount', bold_center)
        worksheet.write(row, 25, 'Tax Amount', bold_center)
        row += 1
        for rec_data in report_data_main[0]:
            partner_name = rec_data['name']
            partner_from = excel_style(row + 1, 1)
            partner_to = excel_style(row + 1, 26)
            worksheet.merge_range('%s:%s'%(partner_from, partner_to), partner_name, bold_center)
            row += 1
            part_unall = 0.00
            part_due_30 = 0.00
            part_due_30_duty = 0.00
            part_due_30_vat = 0.00
            part_due_60 = 0.00
            part_due_60_duty = 0.00
            part_due_60_vat = 0.00
            part_due_90 = 0.00
            part_due_90_duty = 0.00
            part_due_90_vat = 0.00
            part_due_120 = 0.00
            part_due_120_duty = 0.00
            part_due_120_vat = 0.00
            part_due_180 = 0.00
            part_due_180_duty = 0.00
            part_due_180_vat = 0.00
            part_due_365 = 0.00
            part_due_365_duty = 0.00
            part_due_365_vat = 0.00
            part_due_365_plus = 0.00
            part_due_365_p_duty = 0.00
            part_due_365_p_vat = 0.00
            part_total = 0.00
            part_total_duty = 0.00
            part_total_vat = 0.00
            for line_data in rec_data['child_lines']:
                worksheet.write(row, 0, line_data.get('move'))
                try:
                    worksheet.write_datetime(row, 1, datetime.strptime(line_data.get('date'), "%d/%m/%Y"), date_format)
                except:
                    worksheet.write_datetime(row, 1, datetime.strptime(line_data.get('date'), "%Y-%m-%d"), date_format)
                worksheet.write(row, 2, line_data.get('jrnl'))
                worksheet.write(row, 3, line_data.get('acc_code'))
                line_total = 0.00
                if line_data.get('paid_amount'):
                    worksheet.write_number(row, 4, line_data.get('paid_amount'), no_format)
                    part_unall += line_data.get('paid_amount')
                    line_total += line_data.get('paid_amount')
                else:
                    worksheet.write_number(row, 4, 0.00, no_format)
                if line_data.get('period7'):
                    worksheet.write_number(row, 5, line_data.get('amount'), no_format)
                    worksheet.write_number(row, 6, line_data.get('duty_amount'), no_format)
                    worksheet.write_number(row, 7, line_data.get('vat_amount'), no_format)
                    part_due_30_duty += line_data.get('duty_amount')
                    part_due_30_vat += line_data.get('vat_amount')
                    part_due_30 += line_data.get('amount')
                    line_total += line_data.get('amount')
                else:
                    worksheet.write_number(row, 5, 0.00, no_format)
                    worksheet.write_number(row, 6, 0.00, no_format)
                    worksheet.write_number(row, 7, 0.00, no_format)
                if line_data.get('period6'):
                    worksheet.write_number(row, 8, line_data.get('amount'), no_format)
                    worksheet.write_number(row, 9, line_data.get('duty_amount'), no_format)
                    worksheet.write_number(row, 10, line_data.get('vat_amount'), no_format)
                    part_due_60_duty += line_data.get('duty_amount')
                    part_due_60_vat += line_data.get('vat_amount')
                    part_due_60 += line_data.get('amount')
                    line_total += line_data.get('amount')
                else:
                    worksheet.write_number(row, 8, 0.00, no_format)
                    worksheet.write_number(row, 9, 0.00, no_format)
                    worksheet.write_number(row, 10, 0.00, no_format)
                if line_data.get('period5'):
                    worksheet.write_number(row, 11, line_data.get('amount'), no_format)
                    worksheet.write_number(row, 12, line_data.get('duty_amount'), no_format)
                    worksheet.write_number(row, 13, line_data.get('vat_amount'), no_format)
                    part_due_90_duty += line_data.get('duty_amount')
                    part_due_90_vat += line_data.get('vat_amount')
                    part_due_90 += line_data.get('amount')
                    line_total += line_data.get('amount')
                else:
                    worksheet.write_number(row, 11, 0.00, no_format)
                    worksheet.write_number(row, 12, 0.00, no_format)
                    worksheet.write_number(row, 13, 0.00, no_format)
                if line_data.get('period4'):
                    worksheet.write_number(row, 14, line_data.get('amount'), no_format)
                    worksheet.write_number(row, 15, line_data.get('duty_amount'), no_format)
                    worksheet.write_number(row, 16, line_data.get('vat_amount'), no_format)
                    part_due_120_duty += line_data.get('duty_amount')
                    part_due_120_vat += line_data.get('vat_amount')
                    part_due_120 += line_data.get('amount')
                    line_total += line_data.get('amount')
                else:
                    worksheet.write_number(row, 14, 0.00, no_format)
                    worksheet.write_number(row, 15, 0.00, no_format)
                    worksheet.write_number(row, 16, 0.00, no_format)
                if line_data.get('period3'):
                    worksheet.write_number(row, 17, line_data.get('amount'), no_format)
                    worksheet.write_number(row, 18, line_data.get('duty_amount'), no_format)
                    worksheet.write_number(row, 19, line_data.get('vat_amount'), no_format)
                    part_due_180_duty += line_data.get('duty_amount')
                    part_due_180_vat += line_data.get('vat_amount')
                    part_due_180 += line_data.get('amount')
                    line_total += line_data.get('amount')
                else:
                    worksheet.write_number(row, 17, 0.00, no_format)
                    worksheet.write_number(row, 18, 0.00, no_format)
                    worksheet.write_number(row, 19, 0.00, no_format)
                if line_data.get('period2'):
                    worksheet.write_number(row, 20, line_data.get('amount'), no_format)
                    worksheet.write_number(row, 21, line_data.get('duty_amount'), no_format)
                    worksheet.write_number(row, 22, line_data.get('vat_amount'), no_format)
                    part_due_365_duty += line_data.get('duty_amount')
                    part_due_365_vat += line_data.get('vat_amount')
                    part_due_365 += line_data.get('amount')
                    line_total += line_data.get('amount')
                else:
                    worksheet.write_number(row, 20, 0.00, no_format)
                    worksheet.write_number(row, 21, 0.00, no_format)
                    worksheet.write_number(row, 22, 0.00, no_format)
                if line_data.get('period1'):
                    worksheet.write_number(row, 23, line_data.get('amount'), no_format)
                    worksheet.write_number(row, 24, line_data.get('duty_amount'), no_format)
                    worksheet.write_number(row, 25, line_data.get('vat_amount'), no_format)
                    part_due_365_p_duty += line_data.get('duty_amount')
                    part_due_365_p_vat += line_data.get('vat_amount')
                    part_due_365_plus += line_data.get('amount')
                    line_total += line_data.get('amount')
                else:
                    worksheet.write_number(row, 23, 0.00, no_format)
                    worksheet.write_number(row, 24, 0.00, no_format)
                    worksheet.write_number(row, 25, 0.00, no_format)
                part_total += line_total
                row += 1
            worksheet.write(row, 0, "Total", bold)
            worksheet.write_number(row, 4, part_unall, normal_num_bold)
            worksheet.write_number(row, 5, part_due_30, normal_num_bold)
            worksheet.write_number(row, 6, part_due_30_duty, normal_num_bold)
            worksheet.write_number(row, 7, part_due_30_vat, normal_num_bold)
            worksheet.write_number(row, 8, part_due_60, normal_num_bold)
            worksheet.write_number(row, 9, part_due_60_duty, normal_num_bold)
            worksheet.write_number(row, 10, part_due_60_vat, normal_num_bold)
            worksheet.write_number(row, 11, part_due_90, normal_num_bold)
            worksheet.write_number(row, 12, part_due_90_duty, normal_num_bold)
            worksheet.write_number(row, 13, part_due_90_vat, normal_num_bold)
            worksheet.write_number(row, 14, part_due_120, normal_num_bold)
            worksheet.write_number(row, 15, part_due_120_duty, normal_num_bold)
            worksheet.write_number(row, 16, part_due_120_vat, normal_num_bold)
            worksheet.write_number(row, 17, part_due_180, normal_num_bold)
            worksheet.write_number(row, 18, part_due_180_duty, normal_num_bold)
            worksheet.write_number(row, 19, part_due_180_vat, normal_num_bold)
            worksheet.write_number(row, 20, part_due_365, normal_num_bold)
            worksheet.write_number(row, 21, part_due_365_duty, normal_num_bold)
            worksheet.write_number(row, 22, part_due_365_vat, normal_num_bold)
            worksheet.write_number(row, 23, part_due_365_plus, normal_num_bold)
            worksheet.write_number(row, 24, part_due_365_p_duty, normal_num_bold)
            worksheet.write_number(row, 25, part_due_365_p_vat, normal_num_bold)
            row += 1
        workbook.close()
        output.seek(0)
        response.stream.write(output.read())
        output.close()
