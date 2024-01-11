import time
from odoo import fields, models, api, _

import io
import json
from odoo.exceptions import AccessError, UserError, AccessDenied

try:
    from odoo.tools.misc import xlsxwriter
except ImportError:
    import xlsxwriter

from datetime import datetime

class TrialView(models.TransientModel):
    _inherit = "account.common.report"
    _name = 'account.trial.balance'

    journal_ids = fields.Many2many('account.journal',

                                   string='Journals', required=True,
                                   default=[])
    display_account = fields.Selection(
        [('all', 'All'), ('movement', 'With movements'),
         ('not_zero', 'With balance is not equal to 0')],
        string='Display Accounts', required=True, default='movement')

    @api.model
    def view_report(self, option):
        r = self.env['account.trial.balance'].search([('id', '=', option[0])])

        data = {
            'display_account': r.display_account,
            'model':self,
            'journals': r.journal_ids,
            'target_move': r.target_move,

        }
        if r.date_from:
            data.update({
                'date_from':r.date_from,
            })
        if r.date_to:
            data.update({
                'date_to':r.date_to,
            })

        filters = r.get_filter(option)
        records = r._get_report_values(data)
        currency = r._get_currency()

        return {
            'name': "Trial Balance",
            'type': 'ir.actions.client',
            'tag': 't_b',
            'filters': filters,
            'report_lines': records['Accounts'],
            'debit_total': records['debit_total'],
            'credit_total': records['credit_total'],
            'op_debit_total': records['op_debit_total'],
            'op_credit_total': records['op_credit_total'],
            'cl_debit_total': records['cl_debit_total'],
            'cl_credit_total': records['cl_credit_total'],
            'currency': currency,
        }

    def get_filter(self, option):
        data = self.get_filter_data(option)
        filters = {}
        if data.get('journal_ids'):
            filters['journals'] = self.env['account.journal'].browse(data.get('journal_ids')).mapped('code')
        else:
            filters['journals'] = ['All']
        if data.get('target_move'):
            filters['target_move'] = data.get('target_move')
        if data.get('date_from'):
            filters['date_from'] = data.get('date_from')
        if data.get('date_to'):
            filters['date_to'] = data.get('date_to')

        filters['company_id'] = ''
        filters['journals_list'] = data.get('journals_list')
        filters['company_name'] = data.get('company_name')
        filters['target_move'] = data.get('target_move').capitalize()

        return filters

    def get_filter_data(self, option):
        r = self.env['account.trial.balance'].search([('id', '=', option[0])])
        default_filters = {}
        company_id = self.env.company
        company_domain = [('company_id', '=', company_id.id)]
        journals = r.journal_ids if r.journal_ids else self.env['account.journal'].search(company_domain)

        filter_dict = {
            'journal_ids': r.journal_ids.ids,
            'company_id': company_id.id,
            'date_from': r.date_from,
            'date_to': r.date_to,
            'target_move': r.target_move,
            'journals_list': [(j.id, j.name, j.code) for j in journals],
            'company_name': company_id and company_id.name,
        }
        filter_dict.update(default_filters)
        return filter_dict

    def _get_report_values(self, data):
        docs = data['model']
        display_account = data['display_account']
        journals = data['journals']
        accounts = self.env['account.account'].search([])
        if not accounts:
            raise UserError(_("No Accounts Found! Please Add One"))
        account_res = self._get_accounts(accounts, display_account, data)
        debit_total = 0
        debit_total = sum(x['debit'] for x in account_res)
        credit_total = sum(x['credit'] for x in account_res)
        op_debit_total = sum(x['Init_balance']['debit'] for x in account_res)
        op_credit_total = sum(x['Init_balance']['credit'] for x in account_res)
        cl_debit_total = sum(x['closing_balance']['debit'] for x in account_res)
        cl_credit_total = sum(x['closing_balance']['credit'] for x in account_res)
        return {
            'doc_ids': self.ids,
            'debit_total': debit_total,
            'credit_total': credit_total,
            'op_debit_total': op_debit_total,
            'op_credit_total': op_credit_total,
            'cl_debit_total': cl_debit_total,
            'cl_credit_total': cl_credit_total,
            'docs': docs,
            'time': time,
            'Accounts': account_res,
        }

    @api.model
    def create(self, vals):
        vals['target_move'] = 'posted'
        res = super(TrialView, self).create(vals)
        return res

    def write(self, vals):
        if vals.get('target_move'):
            vals.update({'target_move': vals.get('target_move').lower()})
        if vals.get('journal_ids'):
            vals.update({'journal_ids': [(6, 0, vals.get('journal_ids'))]})
        if vals.get('journal_ids') == []:
            vals.update({'journal_ids': [(5,)]})
        res = super(TrialView, self).write(vals)
        return res

    def _get_accounts(self, accounts, display_account, data):

        account_result = {}
        # Prepare sql query base on selected parameters from wizard
        tables, where_clause, where_params = self.env['account.move.line']._query_get()
        tables = tables.replace('"', '')
        if not tables:
            tables = 'account_move_line'
        wheres = [""]
        if where_clause.strip():
            wheres.append(where_clause.strip())
        filters = " AND ".join(wheres)
        if data['target_move'] == 'posted':
            filters += " AND account_move_line__move_id.state = 'posted'"
        else:
            filters += " AND account_move_line__move_id.state in ('draft','posted')"
        if data.get('date_from'):
            filters += " AND account_move_line.date >= '%s'" % data.get('date_from')
        if data.get('date_to'):
            filters += " AND account_move_line.date <= '%s'" % data.get('date_to')

        if data['journals']:
            filters += ' AND jrnl.id IN %s' % str(tuple(data['journals'].ids) + tuple([0]))
        tables += 'JOIN account_journal jrnl ON (account_move_line.journal_id=jrnl.id)'
        # compute the balance, debit and credit for the provided accounts
        request = (
                    "SELECT account_id AS id, SUM(debit) AS debit, SUM(credit) AS credit, (SUM(debit) - SUM(credit)) AS balance" + \
                    " FROM " + tables + " WHERE account_id IN %s " + filters + " GROUP BY account_id")
        params = (tuple(accounts.ids),) + tuple(where_params)
        self.env.cr.execute(request, params)
        for row in self.env.cr.dictfetchall():
            account_result[row.pop('id')] = row

        account_res = []
        for account in accounts:
            res = dict((fn, 0.0) for fn in ['credit', 'debit', 'balance'])
            currency = account.currency_id and account.currency_id or account.company_id.currency_id
            res['code'] = account.code
            res['name'] = account.name
            res['id'] = account.id
            initial_balance = self.get_init_bal(account, display_account, data)
            res['Init_balance'] = initial_balance
            net_debit = initial_balance['debit']
            net_credit = initial_balance['credit']
            if account.id in account_result:
                res['debit'] = account_result[account.id].get('debit')
                res['credit'] = account_result[account.id].get('credit')
                res['balance'] = account_result[account.id].get('balance')
                net_debit += account_result[account.id].get('debit', 0.00)
                net_credit += account_result[account.id].get('credit', 0.00)
            if display_account == 'all':
                account_res.append(res)
            if display_account == 'not_zero' and not currency.is_zero(
                    res['balance']):
                account_res.append(res)
            net_balance = net_debit - net_credit
            res['closing_balance'] = {
                'debit': net_balance > 0.00 and net_balance or 0.00,
                'credit': net_balance < 0.00 and abs(net_balance) or 0.00,
                }
            # if display_account == 'movement' and (
            #         not currency.is_zero(res['closing_balance']['debit']) or not currency.is_zero(
            #         res['closing_balance']['credit'])):
            account_res.append(res)
        return account_res

    def get_init_bal(self, account, display_account, data):
        row = {
            'id': account.id,
            'debit': 0.00,
            'credit': 0.00,
            'balance': 0.00
            }
        if data.get('date_from'):
            tables, where_clause, where_params = self.env[
                'account.move.line']._query_get()
            tables = tables.replace('"', '')
            if not tables:
                tables = 'account_move_line'
            wheres = [""]
            if where_clause.strip():
                wheres.append(where_clause.strip())
            filters = " AND ".join(wheres)
            if data['target_move'] == 'posted':
                filters += " AND account_move_line__move_id.state = 'posted'"
            else:
                filters += " AND account_move_line__move_id.state in ('draft','posted')"
            if data.get('date_from'):
                filters += " AND account_move_line.date < '%s'" % data.get('date_from')

            if data['journals']:
                filters += ' AND jrnl.id IN %s' % str(tuple(data['journals'].ids) + tuple([0]))
            tables += ' JOIN account_journal jrnl ON (account_move_line.journal_id=jrnl.id)'
            tables += ' LEFT JOIN account_account aa on aa.id = account_move_line.account_id'
            tables += ' LEFT JOIN account_account_type aat ON (aat.id=aa.user_type_id)'

            # compute the balance, debit and credit for the provided accounts
            request = (
                    "SELECT account_id AS id, SUM(debit) AS debit, SUM(credit) AS credit, (SUM(debit) - SUM(credit)) AS balance" + \
                    " FROM " + tables + " WHERE account_id = %s and aat.internal_group not in ('income', 'expense')" % account.id + filters + " GROUP BY account_id")
            params = tuple(where_params)
            print (request)
            print ("\n\n\n", params)
            self.env.cr.execute(request, params)
            result = self.env.cr.dictfetchall()
            if result:
                debit = result[0]['debit']
                credit = result[0]['credit']
                balance = debit - credit
                if balance > 0.00:
                    row['debit'] = balance
                if balance < 0.00:
                    row['credit'] = abs(balance)
                row['balance'] = result[0]['balance']
            if account.ret_earning_account:
                request = (
                        "SELECT SUM(debit) AS debit, SUM(credit) AS credit, (SUM(debit) - SUM(credit)) AS balance" + \
                        " FROM " + tables + " WHERE aat.internal_group in ('income', 'expense')" + filters)
                params = tuple(where_params)
                self.env.cr.execute(request, params)
                pl_result = self.env.cr.dictfetchall()
                if pl_result:
                    debit = pl_result[0]['debit'] or 0.00
                    credit = pl_result[0]['credit'] or 0.00
                    acc_debit = row['debit'] or 0.00
                    acc_credit = row['credit'] or 0.00
                    pl_balance = debit - credit
                    if pl_balance < 0.00:
                        acc_credit += abs(pl_balance)
                    if pl_balance > 0.00:
                        acc_debit += pl_balance
                    acc_balance = acc_debit - acc_credit
                    if acc_balance > 0.00:
                        row['debit'] = acc_balance
                    if acc_balance < 0.00:
                        row['credit'] = abs(acc_balance)
                    row['balance'] = acc_balance
        return row

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
                          lang]
        return currency_array

    def get_dynamic_xlsx_report(self, data, response ,report_data, dfr_data):
        report_data_main = json.loads(report_data)
        output = io.BytesIO()
        total = json.loads(dfr_data)
        filters = json.loads(data)
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Trial Balance')
        ##FORMATS##
        heading_format = workbook.add_format({'align': 'center',
                                              'valign': 'vcenter',
                                              'bold': True, 'size': 18})
        sub_heading_format = workbook.add_format({'align': 'center',
                                                  'valign': 'vcenter',
                                                  'bold': True, 'size': 14})
        bold = workbook.add_format({'bold': True})
        bold_center = workbook.add_format({'bold': True, 'valign': 'vcenter', 'bg_color': '#b5b5b5'})
        bold_center_bg = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#b5b5b5'})
        bold_center_center = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#b5b5b5'})
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
        no_format = workbook.add_format({'num_format': '#,##0.00'})
        normal_num_bold = workbook.add_format({'bold': True, 'num_format': '#,##0.00'})
        normal_num_bold_bg = workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'bg_color': '#b5b5b5'})
        ##FORMATS ENDS##
        worksheet.merge_range('A1:H1', "Trial Balance", sub_heading_format)
        row = 2
        if filters.get('date_from', False):
            worksheet.write(row, 0, 'From', bold)
            worksheet.write_datetime(row, 1, datetime.strptime(filters.get('date_from'), "%Y-%m-%d"), date_format)
            row += 1
        if filters.get('date_to', False):
            worksheet.write(row, 0, 'To', bold)
            worksheet.write_datetime(row, 1, datetime.strptime(filters.get('date_to'), "%Y-%m-%d"), date_format)
            row += 1
        worksheet.write(row, 0, 'Journals', bold)
        worksheet.write(row, 1, ', '.join([ lt or '' for lt in filters['journals']]))
        row += 1
        worksheet.write(row, 0, 'Target Move', bold)
        worksheet.write(row, 1, filters.get('target_move'))
        row += 2
        worksheet.set_column('A:A', 35)
        worksheet.set_column('B:B', 25)
        worksheet.set_column('C:C', 50)
        worksheet.set_column('D:D', 25)
        worksheet.set_column('E:E', 25)
        worksheet.set_column('F:F', 25)
        worksheet.set_column('G:G', 25)
        worksheet.set_column('H:H', 25)
        
        
        worksheet.merge_range('C%s:D%s'%(row + 1, row + 1), "Opening Balance", bold_center_bg)
        worksheet.merge_range('E%s:F%s'%(row + 1, row + 1), "Current Transaction", bold_center_bg)
        worksheet.merge_range('G%s:H%s'%(row + 1, row + 1), "Closing Balance", bold_center_bg)
        row += 1
        worksheet.write(row, 0, 'Account Code', bold_center)
        worksheet.write(row, 1, 'Account Name', bold_center)
        worksheet.write(row, 2, 'Opening Debit', bold_center_bg)
        worksheet.write(row, 3, 'Opening Credit', bold_center_bg)
        worksheet.write(row, 4, 'Debit', bold_center_bg)
        worksheet.write(row, 5, 'Credit', bold_center_bg)
        worksheet.write(row, 6, 'Closing Debit', bold_center_bg)
        worksheet.write(row, 7, 'Closing Credit', bold_center_bg)
        row += 1
        for rec_data in report_data_main:
            worksheet.write(row, 0, rec_data['code'])
            worksheet.write(row, 1, rec_data['name'])
            worksheet.write_number(row, 2, rec_data['Init_balance']['debit'], no_format)
            worksheet.write_number(row, 3, rec_data['Init_balance']['credit'], no_format)
            worksheet.write_number(row, 4, rec_data['debit'], no_format)
            worksheet.write_number(row, 5, rec_data['credit'], no_format)
            worksheet.write_number(row, 6, rec_data['closing_balance']['debit'], no_format)
            worksheet.write_number(row, 7, rec_data['closing_balance']['credit'], no_format)
            row += 1
        worksheet.write(row, 1, 'Total', bold_center)
        worksheet.write_number(row, 2, total['op_debit_total'], normal_num_bold_bg)
        worksheet.write_number(row, 3, total['op_credit_total'], normal_num_bold_bg)
        worksheet.write_number(row, 4, total['debit_total'], normal_num_bold_bg)
        worksheet.write_number(row, 5, total['credit_total'], normal_num_bold_bg)
        worksheet.write_number(row, 6, total['cl_debit_total'], normal_num_bold_bg)
        worksheet.write_number(row, 7, total['cl_credit_total'], normal_num_bold_bg)
        workbook.close()
        output.seek(0)
        response.stream.write(output.read())
        output.close()
