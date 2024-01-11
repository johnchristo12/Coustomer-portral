import time
from odoo import fields, models, api, _

import io
import json
from odoo.exceptions import AccessError, UserError, AccessDenied
from odoo.http import request
from datetime import datetime
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

class BalanceSheetView(models.TransientModel):
    _name = 'dynamic.balance.sheet.report'

    company_id = fields.Many2one('res.company', required=True,
                                 default=lambda self: self.env.company)
    journal_ids = fields.Many2many('account.journal',
                                   string='Journals', required=True,
                                   default=[])
    account_ids = fields.Many2many("account.account", string="Accounts")
    account_tag_ids = fields.Many2many("account.account.tag",
                                       string="Account Tags")
    analytic_ids = fields.Many2many(
        "account.analytic.account", string="Analytic Accounts")
    analytic_tag_ids = fields.Many2many("account.analytic.tag",
                                        string="Analytic Tags")
    display_account = fields.Selection(
        [('all', 'All'), ('movement', 'With movements'),
         ('not_zero', 'With balance is not equal to 0')],
        string='Display Accounts', required=True, default='movement')
    target_move = fields.Selection(
        [('all', 'All'), ('posted', 'Posted')],
        string='Target Move', required=True, default='posted')
    date_from = fields.Date(string="Start date")
    date_to = fields.Date(string="End date")
    date_from_comp = fields.Date(string="Comp. Start date")
    date_to_comp = fields.Date(string="Comp. End date")

    def get_fin_line_details(self, parent, intend, data):
        lines = []
        parent_intend = intend
        vals = {
            'name': parent.name,
            'debit': 0.00,
            'credit': 0.00,
            'balance': 0.00,
            'debit_comp': 0.00,
            'credit_comp': 0.00,
            'balance_comp': 0.00,
            'intend': intend,
            'child_lines': [],
            'hide': parent.hide_heading,
            'heading': True
            }
        child_lines = []
        if parent.type == 'sum':
            vals['view'] = True
            intend += 5
            child_objs = self.env['account.financial.report'].search([('parent_id', '=', parent.id)], order="sequence")
            for child in child_objs:
                child_vals, line_debit, line_credit, line_balance, line_debit_comp, line_credit_comp, line_balance_comp = self.get_fin_line_details(child, intend, data)
                child_lines.extend(child_vals)
                vals['debit'] += line_debit
                vals['credit'] += line_credit
                vals['balance'] += line_balance
                vals['debit_comp'] += line_debit_comp
                vals['credit_comp'] += line_credit_comp
                vals['balance_comp'] += line_balance_comp
        elif parent.type in ['account_type', 'accounts']:
            if parent.type == 'account_type':
                account_type_ids = parent.account_type_ids.ids
                account_ids = self.env['account.account'].search([('user_type_id', 'in', account_type_ids)])
            else:
                account_ids = parent.account_ids
            if account_ids:
                sign = 1
                if parent.sign == '-1':
                    sign = -1
                report_value = self._get_report_values(data, account_ids, sign)
                vals['child_lines'].extend(report_value['account_data'])
                vals['debit'] += report_value['debit_total']
                vals['credit'] += report_value['credit_total']
                vals['balance'] += report_value['debit_balance']
                vals['debit_comp'] += report_value.get('debit_total_comp', 0.00)
                vals['credit_comp'] += report_value.get('credit_total_comp', 0.00)
                vals['balance_comp'] += report_value.get('debit_balance_comp', 0.00)
        lines.append(vals)
        lines.extend(child_lines)
        total_lines = {
            'name': "Total %s"%(parent.name),
            'debit': vals['debit'],
            'credit': vals['credit'],
            'balance': vals['balance'],
            'debit_comp': vals['debit_comp'],
            'credit_comp': vals['credit_comp'],
            'balance_comp': vals['balance_comp'],
            'intend': parent_intend,
            'child_lines': [],
            'hide': False,
            'heading': False
            }
        lines.append(total_lines)
        return lines, vals['debit'], vals['credit'], vals['balance'], vals['debit_comp'], vals['credit_comp'], vals['balance_comp']

    def _get_report_values(self, data, account_ids, sign):
        docs = data['model']
        account_res, debit_total, credit_total, balance_total, debit_total_comp, credit_total_comp, balance_total_comp, has_comp = self._get_accounts(account_ids, data, sign)
        currency = self.env.company.currency_id
        return {
            'debit_total': debit_total,
            'credit_total': credit_total,
            'debit_balance': balance_total,
            'debit_total_comp': debit_total_comp,
            'credit_total_comp': credit_total_comp,
            'debit_balance_comp': balance_total_comp,
            'account_data': account_res,
            'has_comp': has_comp
        }
        
    
    def _get_accounts(self, accounts, data, sign):
        cr = self.env.cr
        MoveLine = self.env['account.move.line']
        move_lines = {x: [] for x in accounts.ids}
        currency_id = self.env.company.currency_id
    
        # Prepare sql query base on selected parameters from wizard
        tables, where_clause, where_params = MoveLine._query_get()
        wheres = [""]
        if where_clause.strip():
            wheres.append(where_clause.strip())
        final_filters = " AND ".join(wheres)
        final_filters = final_filters.replace('account_move_line__move_id',
                                              'm').replace(
            'account_move_line', 'l')
        new_final_filter = final_filters
        comp_final_filter = final_filters
    
        if data['target_move'] == 'posted':
            new_final_filter += " AND m.state = 'posted'"
            comp_final_filter += " AND m.state = 'posted'"
        else:
            new_final_filter += " AND m.state in ('draft','posted')"
            comp_final_filter += " AND m.state in ('draft','posted')"
    
        if data.get('date_from'):
            new_final_filter += " AND l.date >= '%s'" % data.get('date_from')
        if data.get('date_to'):
            new_final_filter += " AND l.date <= '%s'" % data.get('date_to')
        
        has_comp = False
        if data.get('date_from_comp'):
            comp_final_filter += " AND l.date >= '%s'" % data.get('date_from_comp')
            has_comp = True
        if data.get('date_to_comp'):
            comp_final_filter += " AND l.date <= '%s'" % data.get('date_to_comp')
            has_comp = True
    
        if data['journals']:
            new_final_filter += ' AND j.id IN %s' % str(
                tuple(data['journals'].ids) + tuple([0]))
            comp_final_filter += ' AND j.id IN %s' % str(
                tuple(data['journals'].ids) + tuple([0]))
    
        RET_WHERE = "WHERE aat.internal_group in ('income', 'expense')"
        if data.get('accounts'):
            WHERE = "WHERE l.account_id IN %s" % str(
                tuple(data.get('accounts').ids) + tuple([0]))
            filter_accounts = data['accounts']
        else:
            WHERE = "WHERE l.account_id IN %s"
            filter_accounts = accounts
    
        if data['analytics']:
            WHERE += ' AND anl.id IN %s' % str(
                tuple(data.get('analytics').ids) + tuple([0]))
            RET_WHERE += ' AND anl.id IN %s' % str(
                tuple(data.get('analytics').ids) + tuple([0]))
    
        if data['analytic_tags']:
            WHERE += ' AND anltag.account_analytic_tag_id IN %s' % str(
                tuple(data.get('analytic_tags').ids) + tuple([0]))
            RET_WHERE += ' AND anltag.account_analytic_tag_id IN %s' % str(
                tuple(data.get('analytic_tags').ids) + tuple([0]))
    
        # Get move lines base on sql query and Calculate the total balance of move lines
        has_ret_earnings = False
        ret_earning_acc = filter_accounts.filtered(lambda a: a.ret_earning_account)
        if ret_earning_acc:
            if len(ret_earning_acc.ids) > 1:
                raise UserError(_("Retained earnings account cannot be more than 1"))
            has_ret_earnings = True
        if has_comp:
            if has_ret_earnings:
                sql = ("""
                    SELECT
                        COALESCE(SUM(x.credit),0) as credit,
                        COALESCE(SUM(x.debit),0) as debit,
                        COALESCE(SUM(x.balance),0) as balance,
                        COALESCE(SUM(x.credit_comp),0) as credit_comp,
                        COALESCE(SUM(x.debit_comp),0) as debit_comp,
                        COALESCE(SUM(x.balance_comp),0) as balance_comp,
                        x.account_name as account_name,
                        x.id as id
                    FROM
                        (SELECT
                            COALESCE(SUM(l.credit),0) as credit,
                            COALESCE(SUM(l.debit),0) as debit,
                            COALESCE(SUM(l.debit - l.credit),0) * %s as balance,
                            0.00 as credit_comp,
                            0.00 as debit_comp,
                            0.00 as balance_comp,
                            acc.code || ' - ' || acc.name as account_name,
                            acc.id as id
                        FROM
                            account_move_line l
                            JOIN account_move m ON (l.move_id=m.id)
                            LEFT JOIN res_currency c ON (l.currency_id=c.id)
                            LEFT JOIN res_partner p ON (l.partner_id=p.id)
                            LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                            LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                            JOIN account_journal j ON (l.journal_id=j.id)
                            JOIN account_account acc ON (l.account_id = acc.id)"""%(sign)
                            + WHERE + new_final_filter + """ GROUP BY acc.code, acc.name, acc.id
                        UNION
                        SELECT
                            0.00 as credit,
                            0.00 as debit,
                            0.00 as balance,
                            COALESCE(SUM(l.credit),0) as credit_comp,
                            COALESCE(SUM(l.debit),0) as debit_comp,
                            COALESCE(SUM(l.debit - l.credit),0) * %s as balance_comp,
                            acc.code || ' - ' || acc.name as account_name,
                            acc.id as id
                        FROM
                            account_move_line l
                            JOIN account_move m ON (l.move_id=m.id)
                            LEFT JOIN res_currency c ON (l.currency_id=c.id)
                            LEFT JOIN res_partner p ON (l.partner_id=p.id)
                            LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                            LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                            JOIN account_journal j ON (l.journal_id=j.id)
                            JOIN account_account acc ON (l.account_id = acc.id)"""%(sign)
                            + WHERE + comp_final_filter + """ GROUP BY acc.code, acc.name, acc.id
                            
                            
                            
                            
                            
                            
                        UNION
                        SELECT
                            CASE WHEN (SUM(l.debit - l.credit) <= 0) THEN SUM(l.debit - l.credit) * -1 ELSE 0.00 END as credit,
                            CASE WHEN (SUM(l.debit - l.credit) > 0 THEN SUM(l.debit - l.credit) ELSE 0.00 END as debit,
                            COALESCE(SUM(l.debit - l.credit),0) * %s as balance,
                            0.00 as credit_comp,
                            0.00 as debit_comp,
                            0.00 as balance_comp,
                            '%s' || ' - ' || '%s' as account_name,
                            %s as id
                        FROM
                            account_move_line l
                            JOIN account_move m ON (l.move_id=m.id)
                            LEFT JOIN res_currency c ON (l.currency_id=c.id)
                            LEFT JOIN res_partner p ON (l.partner_id=p.id)
                            LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                            LEFT JOIN account_account aa on aa.id = l.account_id
                            LEFT JOIN account_account_type aat ON (aat.id=aa.user_type_id)
                            LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                            JOIN account_journal j ON (l.journal_id=j.id)
                            JOIN account_account acc ON (l.account_id = acc.id)"""%(sign, ret_earning_acc.code, ret_earning_acc.name, ret_earning_acc.id)
                            + RET_WHERE + new_final_filter + """
                        UNION
                        SELECT
                            0.00 as credit,
                            0.00 as debit,
                            0.00 as balance,
                            CASE WHEN (SUM(l.debit - l.credit) <= 0) THEN SUM(l.debit - l.credit) * -1 ELSE 0.00 END as credit_comp,
                            CASE WHEN (SUM(l.debit - l.credit) > 0) THEN SUM(l.debit - l.credit) ELSE 0.00 END as debit_comp,
                            COALESCE(SUM(l.debit - l.credit),0) * %s as balance_comp,
                            '%s' || ' - ' || '%s' as account_name,
                            %s as id
                        FROM
                            account_move_line l
                            JOIN account_move m ON (l.move_id=m.id)
                            LEFT JOIN res_currency c ON (l.currency_id=c.id)
                            LEFT JOIN res_partner p ON (l.partner_id=p.id)
                            LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                            LEFT JOIN account_account aa on aa.id = l.account_id
                            LEFT JOIN account_account_type aat ON (aat.id=aa.user_type_id)
                            LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                            JOIN account_journal j ON (l.journal_id=j.id)
                            JOIN account_account acc ON (l.account_id = acc.id)"""%(sign, ret_earning_acc.code, ret_earning_acc.name, ret_earning_acc.id)
                            + RET_WHERE + comp_final_filter + """ 
                            
                            )x
                    GROUP BY x.account_name, x.id""")
                if data.get('accounts'):
                    params = tuple(where_params) + tuple(where_params) + tuple(where_params) + tuple(where_params)
                else:
                    params = (tuple(accounts.ids),) + tuple(where_params) + (tuple(accounts.ids),) + tuple(where_params) + tuple(where_params) + tuple(where_params)
            else:
                sql = ("""
                    SELECT
                        COALESCE(SUM(x.credit),0) as credit,
                        COALESCE(SUM(x.debit),0) as debit,
                        COALESCE(SUM(x.balance),0) as balance,
                        COALESCE(SUM(x.credit_comp),0) as credit_comp,
                        COALESCE(SUM(x.debit_comp),0) as debit_comp,
                        COALESCE(SUM(x.balance_comp),0) as balance_comp,
                        x.account_name as account_name,
                        x.id as id
                    FROM
                        (SELECT
                            COALESCE(SUM(l.credit),0) as credit,
                            COALESCE(SUM(l.debit),0) as debit,
                            COALESCE(SUM(l.debit - l.credit),0) * %s as balance,
                            0.00 as credit_comp,
                            0.00 as debit_comp,
                            0.00 as balance_comp,
                            acc.code || ' - ' || acc.name as account_name,
                            acc.id as id
                        FROM
                            account_move_line l
                            JOIN account_move m ON (l.move_id=m.id)
                            LEFT JOIN res_currency c ON (l.currency_id=c.id)
                            LEFT JOIN res_partner p ON (l.partner_id=p.id)
                            LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                            LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                            JOIN account_journal j ON (l.journal_id=j.id)
                            JOIN account_account acc ON (l.account_id = acc.id)"""%(sign)
                            + WHERE + new_final_filter + """ GROUP BY acc.code, acc.name, acc.id
                        UNION
                        SELECT
                            0.00 as credit,
                            0.00 as debit,
                            0.00 as balance,
                            COALESCE(SUM(l.credit),0) as credit_comp,
                            COALESCE(SUM(l.debit),0) as debit_comp,
                            COALESCE(SUM(l.debit - l.credit),0) * %s as balance_comp,
                            acc.code || ' - ' || acc.name as account_name,
                            acc.id as id
                        FROM
                            account_move_line l
                            JOIN account_move m ON (l.move_id=m.id)
                            LEFT JOIN res_currency c ON (l.currency_id=c.id)
                            LEFT JOIN res_partner p ON (l.partner_id=p.id)
                            LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                            LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                            JOIN account_journal j ON (l.journal_id=j.id)
                            JOIN account_account acc ON (l.account_id = acc.id)"""%(sign)
                            + WHERE + comp_final_filter + """ GROUP BY acc.code, acc.name, acc.id)x
                    GROUP BY x.account_name, x.id""")
                if data.get('accounts'):
                    params = tuple(where_params) + tuple(where_params)
                else:
                    params = (tuple(accounts.ids),) + tuple(where_params) + (tuple(accounts.ids),) + tuple(where_params)
        else:
            if has_ret_earnings:
                sql = ("""
                    SELECT
                        COALESCE(SUM(x.credit),0) as credit,
                        COALESCE(SUM(x.debit),0) as debit,
                        COALESCE(SUM(x.balance),0) as balance,
                        COALESCE(SUM(x.credit_comp),0) as credit_comp,
                        COALESCE(SUM(x.debit_comp),0) as debit_comp,
                        COALESCE(SUM(x.balance_comp),0) as balance_comp,
                        x.account_name as account_name,
                        x.id as id
                    FROM
                    (SELECT
                        COALESCE(SUM(l.credit),0) as credit,
                        COALESCE(SUM(l.debit),0) as debit,
                        COALESCE(SUM(l.debit - l.credit),0) * %s as balance,
                        0.00 as credit_comp,
                        0.00 as debit_comp,
                        0.00 as balance_comp,
                        acc.code || ' - ' || acc.name as account_name,
                        acc.id as id
                    FROM
                        account_move_line l
                        JOIN account_move m ON (l.move_id=m.id)
                        LEFT JOIN res_currency c ON (l.currency_id=c.id)
                        LEFT JOIN res_partner p ON (l.partner_id=p.id)
                        LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                        LEFT JOIN account_account aa on aa.id = l.account_id
                        LEFT JOIN account_account_type aat ON (aat.id=aa.user_type_id)
                        LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                        JOIN account_journal j ON (l.journal_id=j.id)
                        JOIN account_account acc ON (l.account_id = acc.id)"""%(sign)
                        + WHERE + new_final_filter + """ GROUP BY acc.code, acc.name, acc.id
                    UNION
                    SELECT
                        CASE WHEN SUM(l.debit - l.credit) < 0 THEN SUM(l.debit - l.credit) * -1 ELSE 0 END as credit,
                        CASE WHEN SUM(l.debit - l.credit) > 0 THEN SUM(l.debit - l.credit) ELSE 0 END as debit,
                        COALESCE(SUM(l.debit - l.credit),0) * %s as balance,
                        0.00 as credit_comp,
                        0.00 as debit_comp,
                        0.00 as balance_comp,
                        '%s' || ' - ' || '%s' as account_name,
                        %s as id
                    FROM
                        account_move_line l
                        JOIN account_move m ON (l.move_id=m.id)
                        LEFT JOIN res_currency c ON (l.currency_id=c.id)
                        LEFT JOIN res_partner p ON (l.partner_id=p.id)
                        LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                        LEFT JOIN account_account aa on aa.id = l.account_id
                        LEFT JOIN account_account_type aat ON (aat.id=aa.user_type_id)
                        LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                        JOIN account_journal j ON (l.journal_id=j.id)
                        JOIN account_account acc ON (l.account_id = acc.id)"""%(sign, ret_earning_acc.code, ret_earning_acc.name, ret_earning_acc.id)
                        + RET_WHERE + new_final_filter + """)x group by x.account_name, x.id
                    """)
                if data.get('accounts'):
                    params = tuple(where_params) + tuple(where_params)
                else:
                    params = (tuple(accounts.ids),) + tuple(where_params) + tuple(where_params)
            else:
                sql = ("""
                    SELECT
                        COALESCE(SUM(l.credit),0) as credit,
                        COALESCE(SUM(l.debit),0) as debit,
                        COALESCE(SUM(l.debit - l.credit),0) * %s as balance,
                        0.00 as credit_comp,
                        0.00 as debit_comp,
                        0.00 as balance_comp,
                        acc.code || ' - ' || acc.name as account_name,
                        acc.id as id
                    FROM
                        account_move_line l
                        JOIN account_move m ON (l.move_id=m.id)
                        LEFT JOIN res_currency c ON (l.currency_id=c.id)
                        LEFT JOIN res_partner p ON (l.partner_id=p.id)
                        LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                        LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                        JOIN account_journal j ON (l.journal_id=j.id)
                        JOIN account_account acc ON (l.account_id = acc.id)"""%(sign)
                        + WHERE + new_final_filter + """ GROUP BY acc.code, acc.name, acc.id""")
                if data.get('accounts'):
                    params = tuple(where_params)
                else:
                    params = (tuple(accounts.ids),) + tuple(where_params)
        cr.execute(sql, params)
        account_res = cr.dictfetchall()
        if has_comp:
            if has_ret_earnings:
                bal_sql = ("""
                    SELECT
                        COALESCE(SUM(x.credit),0) as credit,
                        COALESCE(SUM(x.debit),0) as debit,
                        COALESCE(SUM(x.balance),0) as balance,
                        COALESCE(SUM(x.credit_comp),0) as credit_comp,
                        COALESCE(SUM(x.debit_comp),0) as debit_comp,
                        COALESCE(SUM(x.balance_comp),0) as balance_comp
                    FROM
                        (SELECT
                            COALESCE(SUM(l.credit),0) as credit,
                            COALESCE(SUM(l.debit),0) as debit,
                            COALESCE(SUM(l.debit - l.credit),0) * %s as balance,
                            0.00 as credit_comp,
                            0.00 as debit_comp,
                            0.00 as balance_comp
                        FROM
                            account_move_line l
                            JOIN account_move m ON (l.move_id=m.id)
                            LEFT JOIN res_currency c ON (l.currency_id=c.id)
                            LEFT JOIN res_partner p ON (l.partner_id=p.id)
                            LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                            LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                            JOIN account_journal j ON (l.journal_id=j.id)
                            JOIN account_account acc ON (l.account_id = acc.id)"""%(sign)
                            + WHERE + new_final_filter + """ 
                        UNION
                        SELECT
                            0.00 as credit,
                            0.00 as debit,
                            0.00 as balance,
                            COALESCE(SUM(l.credit),0) as credit_comp,
                            COALESCE(SUM(l.debit),0) as debit_comp,
                            COALESCE(SUM(l.debit - l.credit),0) * %s as balance_comp
                        FROM
                            account_move_line l
                            JOIN account_move m ON (l.move_id=m.id)
                            LEFT JOIN res_currency c ON (l.currency_id=c.id)
                            LEFT JOIN res_partner p ON (l.partner_id=p.id)
                            LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                            LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                            JOIN account_journal j ON (l.journal_id=j.id)
                            JOIN account_account acc ON (l.account_id = acc.id)"""%(sign)
                            + WHERE + comp_final_filter + """ 
                        UNION
                        SELECT
                            CASE WHEN (SUM(l.debit - l.credit) <= 0) THEN SUM(l.debit - l.credit) * -1 ELSE 0.00 END as credit,
                            CASE WHEN (SUM(l.debit - l.credit) > 0) THEN SUM(l.debit - l.credit) ELSE 0.00 END as debit,
                            COALESCE(SUM(l.debit - l.credit),0) * %s as balance,
                            0.00 as credit_comp,
                            0.00 as debit_comp,
                            0.00 as balance_comp
                        FROM
                            account_move_line l
                            JOIN account_move m ON (l.move_id=m.id)
                            LEFT JOIN res_currency c ON (l.currency_id=c.id)
                            LEFT JOIN res_partner p ON (l.partner_id=p.id)
                            LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                            LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                            JOIN account_journal j ON (l.journal_id=j.id)
                            JOIN account_account acc ON (l.account_id = acc.id)
                            LEFT JOIN account_account_type aat ON (aat.id=acc.user_type_id)"""%(sign)
                            + RET_WHERE + new_final_filter + """ 
                        UNION
                        SELECT
                            0.00 as credit,
                            0.00 as debit,
                            0.00 as balance,
                            CASE WHEN (SUM(l.debit - l.credit) <= 0) THEN SUM(l.debit - l.credit) * -1 ELSE 0.00 END as credit_comp,
                            CASE WHEN (SUM(l.debit - l.credit) > 0) THEN SUM(l.debit - l.credit) ELSE 0.00 END as debit_comp,
                            COALESCE(SUM(l.debit - l.credit),0) * %s as balance_comp
                        FROM
                            account_move_line l
                            JOIN account_move m ON (l.move_id=m.id)
                            LEFT JOIN res_currency c ON (l.currency_id=c.id)
                            LEFT JOIN res_partner p ON (l.partner_id=p.id)
                            LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                            LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                            JOIN account_journal j ON (l.journal_id=j.id)
                            JOIN account_account acc ON (l.account_id = acc.id)
                            LEFT JOIN account_account_type aat ON (aat.id=acc.user_type_id)"""%(sign)
                            + RET_WHERE + comp_final_filter + """)x""")
                if data.get('accounts'):
                    params = tuple(where_params) + tuple(where_params) + tuple(where_params) + tuple(where_params)
                else:
                    params = (tuple(accounts.ids),) + tuple(where_params) + (tuple(accounts.ids),) + tuple(where_params) + tuple(where_params) + tuple(where_params)
                
                
            else:
                bal_sql = ("""
                    SELECT
                        COALESCE(SUM(x.credit),0) as credit,
                        COALESCE(SUM(x.debit),0) as debit,
                        COALESCE(SUM(x.balance),0) as balance,
                        COALESCE(SUM(x.credit_comp),0) as credit_comp,
                        COALESCE(SUM(x.debit_comp),0) as debit_comp,
                        COALESCE(SUM(x.balance_comp),0) as balance_comp
                    FROM
                        (SELECT
                            COALESCE(SUM(l.credit),0) as credit,
                            COALESCE(SUM(l.debit),0) as debit,
                            COALESCE(SUM(l.debit - l.credit),0) * %s as balance,
                            0.00 as credit_comp,
                            0.00 as debit_comp,
                            0.00 as balance_comp
                        FROM
                            account_move_line l
                            JOIN account_move m ON (l.move_id=m.id)
                            LEFT JOIN res_currency c ON (l.currency_id=c.id)
                            LEFT JOIN res_partner p ON (l.partner_id=p.id)
                            LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                            LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                            JOIN account_journal j ON (l.journal_id=j.id)
                            JOIN account_account acc ON (l.account_id = acc.id)"""%(sign)
                            + WHERE + new_final_filter + """ 
                        UNION
                        SELECT
                            0.00 as credit,
                            0.00 as debit,
                            0.00 as balance,
                            COALESCE(SUM(l.credit),0) as credit_comp,
                            COALESCE(SUM(l.debit),0) as debit_comp,
                            COALESCE(SUM(l.debit - l.credit),0) * %s as balance_comp
                        FROM
                            account_move_line l
                            JOIN account_move m ON (l.move_id=m.id)
                            LEFT JOIN res_currency c ON (l.currency_id=c.id)
                            LEFT JOIN res_partner p ON (l.partner_id=p.id)
                            LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                            LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                            JOIN account_journal j ON (l.journal_id=j.id)
                            JOIN account_account acc ON (l.account_id = acc.id)"""%(sign)
                            + WHERE + comp_final_filter + """ )x""")
                if data.get('accounts'):
                    params = tuple(where_params) + tuple(where_params)
                else:
                    params = (tuple(accounts.ids),) + tuple(where_params) + (tuple(accounts.ids),) + tuple(where_params)
        else:
            if has_ret_earnings:
                bal_sql = ("""SELECT
                    COALESCE(SUM(x.credit),0) as credit,
                    COALESCE(SUM(x.debit),0) as debit,
                    COALESCE(SUM(x.balance),0) as balance,
                    COALESCE(SUM(x.credit_comp),0) as credit_comp,
                    COALESCE(SUM(x.debit_comp),0) as debit_comp,
                    COALESCE(SUM(x.balance_comp),0) as balance_comp
                FROM
                (SELECT
                        COALESCE(SUM(l.credit),0) as credit,
                        COALESCE(SUM(l.debit),0) as debit,
                        COALESCE(SUM(l.debit - l.credit),0) * %s as balance,
                        0.00 as credit_comp,
                        0.00 as debit_comp,
                        0.00 as balance_comp
                    FROM
                        account_move_line l
                        JOIN account_move m ON (l.move_id=m.id)
                        LEFT JOIN res_currency c ON (l.currency_id=c.id)
                        LEFT JOIN res_partner p ON (l.partner_id=p.id)
                        LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                        LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                        JOIN account_journal j ON (l.journal_id=j.id)
                        JOIN account_account acc ON (l.account_id = acc.id)"""%(sign)
                        + WHERE + new_final_filter + """
                UNION
                SELECT
                        CASE WHEN (SUM(l.debit - l.credit) <= 0) THEN SUM(l.debit - l.credit) * -1 ELSE 0.00 END as credit,
                        CASE WHEN (SUM(l.debit - l.credit) > 0) THEN SUM(l.debit - l.credit) ELSE 0.00 END as debit,
                        COALESCE(SUM(l.debit - l.credit),0) * %s as balance,
                        0.00 as credit_comp,
                        0.00 as debit_comp,
                        0.00 as balance_comp
                    FROM
                        account_move_line l
                        JOIN account_move m ON (l.move_id=m.id)
                        LEFT JOIN res_currency c ON (l.currency_id=c.id)
                        LEFT JOIN res_partner p ON (l.partner_id=p.id)
                        LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                        LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                        JOIN account_journal j ON (l.journal_id=j.id)
                        JOIN account_account acc ON (l.account_id = acc.id)
                        LEFT JOIN account_account_type aat ON (aat.id=acc.user_type_id)"""%(sign)
                        + RET_WHERE + new_final_filter + """)x""")
                if data.get('accounts'):
                    params = tuple(where_params) + tuple(where_params)
                else:
                    params = (tuple(accounts.ids),) + tuple(where_params) + tuple(where_params)
            else:
                bal_sql = ("""
                    SELECT
                        COALESCE(SUM(l.credit),0) as credit,
                        COALESCE(SUM(l.debit),0) as debit,
                        COALESCE(SUM(l.debit - l.credit),0) * %s as balance,
                        0.00 as credit_comp,
                        0.00 as debit_comp,
                        0.00 as balance_comp
                    FROM
                        account_move_line l
                        JOIN account_move m ON (l.move_id=m.id)
                        LEFT JOIN res_currency c ON (l.currency_id=c.id)
                        LEFT JOIN res_partner p ON (l.partner_id=p.id)
                        LEFT JOIN account_analytic_account anl ON (l.analytic_account_id=anl.id)
                        LEFT JOIN account_analytic_tag_account_move_line_rel anltag ON (anltag.account_move_line_id = l.id)
                        JOIN account_journal j ON (l.journal_id=j.id)
                        JOIN account_account acc ON (l.account_id = acc.id)"""%(sign)
                        + WHERE + new_final_filter)
                if data.get('accounts'):
                    params = tuple(where_params)
                else:
                    params = (tuple(accounts.ids),) + tuple(where_params)
        cr.execute(bal_sql, params)
        balance_res = cr.dictfetchall()
        if not balance_res:
            return account_res, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, has_comp
        else:
            return account_res, balance_res[0]['debit'], balance_res[0]['credit'], balance_res[0]['balance'], balance_res[0]['debit_comp'], balance_res[0]['credit_comp'], balance_res[0]['balance_comp'], has_comp

    @api.model
    def view_report(self, option, tag, lang):
        r = self.env['dynamic.balance.sheet.report'].search([('id', '=', option[0])])
        filters = self.get_filter(option)
        data = {
            'display_account': r.display_account,
            'model': self,
            'journals': r.journal_ids,
            'target_move': r.target_move,
            'accounts': r.account_ids,
            'account_tags': r.account_tag_ids,
            'analytics': r.analytic_ids,
            'analytic_tags': r.analytic_tag_ids,
        }
        if r.date_from:
            data.update({
                'date_from': r.date_from,
            })
        if r.date_to:
            data.update({
                'date_to': r.date_to,
            })
        has_comp = False
        if r.date_from_comp:
            data.update({
                'date_from_comp': r.date_from_comp,
            })
            has_comp = True
        if r.date_to_comp:
            data.update({
                'date_to_comp': r.date_to_comp,
            })
            has_comp = True
        account_report_id = self.env['account.financial.report'].search([('name', 'ilike', tag)], order="sequence")
        if not account_report_id:
            raise UserError(_("Unable to find the financial report."))
        intend = 0
        report_lines, net_debit, net_credit, net_balance, net_debit_comp, net_credit_comp, net_balance_comp = self.get_fin_line_details(account_report_id, intend, data)
        return {
            'name': tag,
            'type': 'ir.actions.client',
            'tag': tag,
            'filters': filters,
            'debit_total': net_debit,
            'credit_total': net_credit,
            'debit_balance': net_balance,
            'debit_total_comp': net_debit_comp,
            'credit_total_comp': net_credit_comp,
            'debit_balance_comp': net_balance_comp,
            'currency': self.env.company.currency_id.name,
            'factor': self.env.company.currency_id.decimal_places,
            'bs_lines': report_lines,
            'has_comp': has_comp
        }




    def get_filter(self, option):
        data = self.get_filter_data(option)
        filters = {}
        if data.get('journal_ids'):
            filters['journals'] = self.env['account.journal'].browse(data.get('journal_ids')).mapped('code')
        else:
            filters['journals'] = ['All']
        if data.get('account_ids', []):
            filters['accounts'] = self.env['account.account'].browse(data.get('account_ids', [])).mapped('code')
        else:
            filters['accounts'] = ['All']
        if data.get('target_move'):
            filters['target_move'] = data.get('target_move')
        else:
            filters['target_move'] = 'posted'
        if data.get('date_from'):
            filters['date_from'] = data.get('date_from')
        else:
            filters['date_from'] = False
        if data.get('date_to'):
            filters['date_to'] = data.get('date_to')
        else:
            filters['date_to'] = False
        
        if data.get('date_from_comp'):
            filters['date_from_comp'] = data.get('date_from_comp')
        else:
            filters['date_from_comp'] = False
        if data.get('date_to_comp'):
            filters['date_to_comp'] = data.get('date_to_comp')
        else:
            filters['date_to_comp'] = False
        
        if data.get('analytic_ids', []):
            filters['analytics'] = self.env['account.analytic.account'].browse(data.get('analytic_ids', [])).mapped('name')
        else:
            filters['analytics'] = ['All']
        if data.get('account_tag_ids'):
            filters['account_tags'] = self.env['account.account.tag'].browse(data.get('account_tag_ids', [])).mapped('name')
        else:
            filters['account_tags'] = ['All']
        if data.get('analytic_tag_ids', []):
            filters['analytic_tags'] = self.env['account.analytic.tag'].browse(data.get('analytic_tag_ids', [])).mapped('name')
        else:
            filters['analytic_tags'] = ['All']
        filters['company_id'] = ''
        filters['accounts_list'] = data.get('accounts_list')
        filters['journals_list'] = data.get('journals_list')
        filters['analytic_list'] = data.get('analytic_list')
        filters['account_tag_list'] = data.get('account_tag_list')
        filters['analytic_tag_list'] = data.get('analytic_tag_list')
        filters['company_name'] = data.get('company_name')
        filters['target_move'] = data.get('target_move').capitalize()
        return filters
    
    def get_filter_data(self, option):
        r = self.env['dynamic.balance.sheet.report'].search([('id', '=', option[0])])
        default_filters = {}
        company_id = self.env.company
        company_domain = [('company_id', '=', company_id.id)]
        journals = r.journal_ids if r.journal_ids else self.env['account.journal'].search(company_domain)
        analytics = r.analytic_ids if r.analytic_ids else self.env['account.analytic.account'].search(company_domain)
        account_tags = r.account_tag_ids if r.account_tag_ids else self.env['account.account.tag'].search([])
        analytic_tags = r.analytic_tag_ids if r.analytic_tag_ids else self.env['account.analytic.tag'].sudo().search(['|', ('company_id', '=', company_id.id),
                                                                                                                      ('company_id', '=', False)])
        if r.account_tag_ids:
            company_domain.append(('tag_ids', 'in', r.account_tag_ids.ids))
    
        accounts = self.account_ids if self.account_ids else self.env['account.account'].search(company_domain)
        filter_dict = {
            'journal_ids': r.journal_ids.ids,
            'account_ids': r.account_ids.ids,
            'analytic_ids': r.analytic_ids.ids,
            'company_id': company_id.id,
            'date_from': r.date_from,
            'date_to': r.date_to,
            'date_from_comp': r.date_from_comp,
            'date_to_comp': r.date_to_comp,
            'target_move': r.target_move,
            'journals_list': [(j.id, j.name, j.code) for j in journals],
            'accounts_list': [(a.id, a.name) for a in accounts],
            'analytic_list': [(anl.id, anl.name) for anl in analytics],
            'company_name': company_id and company_id.name,
            'analytic_tag_ids': r.analytic_tag_ids.ids,
            'analytic_tag_list': [(anltag.id, anltag.name) for anltag in
                                  analytic_tags],
            'account_tag_ids': r.account_tag_ids.ids,
            'account_tag_list': [(a.id, a.name) for a in account_tags],
        }
        filter_dict.update(default_filters)
        return filter_dict
    
    def get_dynamic_xlsx_report(self, options, response, report_data, dfr_data):
        i_data = str(report_data)
        filters = json.loads(options)
        j_data = dfr_data
        rl_data = json.loads(j_data)
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        ##FORMATS##
        heading_format = workbook.add_format({'align': 'center',
                                              'valign': 'vcenter',
                                              'bold': True, 'size': 18})
        sub_heading_format = workbook.add_format({'align': 'center',
                                                  'valign': 'vcenter',
                                                  'bold': True, 'size': 14})
        bold = workbook.add_format({'bold': True})
        bold_right = workbook.add_format({'bold': True, 'align': 'right'})
        bold_center = workbook.add_format({'bold': True, 'align': 'center'})
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
        no_format_2 = workbook.add_format({'num_format': '#,##0.00', 'align': 'right'})
        no_format_3 = workbook.add_format({'num_format': '#,##0.000', 'align': 'right'})
        no_format_doll = workbook.add_format({'num_format': '[$$-409]#,##0.00', 'align': 'right'})
        no_format_doll_bold = workbook.add_format({'num_format': '[$$-409]#,##0.00', 'align': 'right', 'bold': True})
        normal_num_bold_2 = workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'align': 'right'})
        normal_num_bold_3 = workbook.add_format({'bold': True, 'num_format': '#,##0.000', 'align': 'right'})
        perc_format_bold = workbook.add_format({'bold': True, 'num_format': '0.00%', 'align': 'right'})
        perc_format = workbook.add_format({'num_format': '0.00%', 'align': 'right'})
        bg_red = workbook.add_format({'bg_color': 'red'})
        bg_green = workbook.add_format({'bg_color': 'green'})
        bg_grey = workbook.add_format({'bg_color': 'grey'})
        ##Format Ends##
        worksheet = workbook.add_worksheet(report_data)
        worksheet.merge_range('A1:D1', self.env.company.name, heading_format)
        worksheet.merge_range('A2:D2', report_data, sub_heading_format)
        worksheet.set_column('A:A', 50)
        worksheet.set_column('B:B', 25)
        worksheet.set_column('C:C', 25)
        worksheet.set_column('D:D', 25)
        worksheet.set_column('E:E', 25)
        worksheet.set_column('F:F', 25)
        worksheet.set_column('G:G', 25)
        row = 3
        decimal = self.env.company.currency_id.decimal_places
        if filters.get('date_from', False):
            worksheet.write(row, 0, "From:", bold)
            worksheet.write_datetime(row, 1, datetime.strptime(filters['date_from'], "%Y-%m-%d"), date_format)
            row += 1
        if filters.get('date_to'):
            worksheet.write(row, 0, "To:", bold)
            worksheet.write_datetime(row, 1, datetime.strptime(filters['date_to'], "%Y-%m-%d"), date_format)
            row += 1
        has_comp = False
        if filters.get('date_from_comp', False):
            worksheet.write(row, 0, "Comparison Date From:", bold)
            worksheet.write_datetime(row, 1, datetime.strptime(filters['date_from_comp'], "%Y-%m-%d"), date_format)
            row += 1
            has_comp = True
        if filters.get('date_to_comp'):
            worksheet.write(row, 0, "Comparison Date To:", bold)
            worksheet.write_datetime(row, 1, datetime.strptime(filters['date_to_comp'], "%Y-%m-%d"), date_format)
            row += 1
            has_comp = True
        filter_from_row = excel_style(row + 1, 1)
        filter_to_row = excel_style(row + 1, 4)
        worksheet.merge_range('%s:%s'%(filter_from_row, filter_to_row), '  Accounts: ' + ', '.join(
            [lt or '' for lt in
             filters['accounts']]) + ';  Journals: ' + ', '.join(
            [lt or '' for lt in
             filters['journals']]) + ';  Account Tags: ' + ', '.join(
            [lt or '' for lt in
             filters['account_tags']]) + ';  Analytic Tags: ' + ', '.join(
            [lt or '' for lt in
             filters['analytic_tags']]) + ';  Analytic: ' + ', '.join(
            [at or '' for at in
             filters['analytics']]) + ';  Target Moves: ' + filters.get(
            'target_move').capitalize())
        row += 2
        worksheet.write(row, 0, 'Account', bold)
        worksheet.write(row, 1, 'Debit', bold)
        worksheet.write(row, 2, 'Credit', bold)
        worksheet.write(row, 3, 'Balance', bold)
        if has_comp:
            worksheet.write(row, 4, 'Comp. Debit', bold)
            worksheet.write(row, 5, 'Comp. Credit', bold)
            worksheet.write(row, 6, 'Comp. Balance', bold)
        row += 1
        for line in rl_data:
            indent = line['intend']
            if not line['hide']:
                worksheet.write(row, 0, "%s%s"%("  "*indent, line['name']), bold)
                if not line['heading']:
                    worksheet.write_number(row, 1, line['debit'], eval('normal_num_bold_%s'%(int(decimal))))
                    worksheet.write_number(row, 2, line['credit'], eval('normal_num_bold_%s'%(int(decimal))))
                    worksheet.write_number(row, 3, line['balance'], eval('normal_num_bold_%s'%(int(decimal))))
                    if has_comp:
                        worksheet.write_number(row, 4, line['debit_comp'], eval('normal_num_bold_%s'%(int(decimal))))
                        worksheet.write_number(row, 5, line['credit_comp'], eval('normal_num_bold_%s'%(int(decimal))))
                        worksheet.write_number(row, 6, line['balance_comp'], eval('normal_num_bold_%s'%(int(decimal))))
                row += 1
            for child_line in line['child_lines']:
                worksheet.write(row, 0, "%s%s"%("  " * (indent + 5), child_line['account_name']))
                worksheet.write_number(row, 1, child_line['debit'], eval('no_format_%s'%(int(decimal))))
                worksheet.write_number(row, 2, child_line['credit'], eval('no_format_%s'%(int(decimal))))
                worksheet.write_number(row, 3, child_line['balance'], eval('no_format_%s'%(int(decimal))))
                if has_comp:
                    worksheet.write_number(row, 4, child_line['debit_comp'], eval('no_format_%s'%(int(decimal))))
                    worksheet.write_number(row, 5, child_line['credit_comp'], eval('no_format_%s'%(int(decimal))))
                    worksheet.write_number(row, 6, child_line['balance_comp'], eval('no_format_%s'%(int(decimal))))
                row += 1
        workbook.close()
        output.seek(0)
        response.stream.write(output.read())
        output.close()


