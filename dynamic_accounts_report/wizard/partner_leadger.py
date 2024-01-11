import time
from odoo import fields, models, api, _

import io
import json
from odoo.exceptions import AccessError, UserError, AccessDenied

try:
    from odoo.tools.misc import xlsxwriter
except ImportError:
    import xlsxwriter


class PartnerView(models.TransientModel):
    _inherit = "account.common.report"
    _name = 'account.partner.ledger'

    journal_ids = fields.Many2many('account.journal',
                                   string='Journals', required=True,
                                   default=[])
    account_ids = fields.Many2many(
        "account.account",
        string="Accounts", check_company=True,
    )

    display_account = fields.Selection(
        [('all', 'All'), ('movement', 'With movements'),
         ('not_zero', 'With balance is not equal to 0')],
        string='Display Accounts', required=True, default='movement')

    partner_ids = fields.Many2many('res.partner', string='Partner')
    partner_category_ids = fields.Many2many('res.partner.category',
                                            string='Partner tags')
    reconciled = fields.Selection([
        ('unreconciled', 'Unreconciled Only')],
        string='Reconcile Type', default='unreconciled')

    account_type_ids = fields.Many2many('account.account.type',string='Account Type',
                                        domain=[('type', 'in', ('receivable', 'payable'))])

    @api.model
    def create(self, vals):
        account_type_id = self.env['account.account.type'].search([('type', '=', 'receivable')])[0]
        vals['account_type_ids'] = account_type_id[0]
        res = super(PartnerView, self).create(vals)
        return res

    def get_where_condition(self, opening=False, detail=False):
        tables, where_clause, where_params = self.env['account.move.line']._query_get()
        where_clause = where_clause.replace("account_move_line__move_id", "am")
        where_clause = where_clause.replace("account_move_line", "aml")
        where_clause += " AND aat.type in ('receivable', 'payable') AND aml.partner_id is not null"
        if self.journal_ids:
            journal_ids = self.journal_ids.ids
            if len(journal_ids) == 1:
                where_clause += " AND am.journal_id = %s"%(journal_ids[0])
            else:
                where_clause += " AND am.journal_id in %s"%(tuple(journal_ids),)
        if self.account_ids:
            account_ids = self.account_ids.ids
            if len(account_ids) == 1:
                where_clause += " AND aml.account_id = %s"%(account_ids[0])
            else:
                where_clause += " AND aml.account_id in %s"%(tuple(account_ids),)
        if self.partner_ids:
            partner_ids = self.partner_ids.ids
            if len(partner_ids) == 1:
                where_clause += " AND aml.partner_id = %s"%(partner_ids[0])
            else:
                where_clause += " AND aml.partner_id in %s"%(tuple(partner_ids),)
        if self.partner_category_ids:
            partner_category_ids = self.partner_category_ids.ids
            if len(partner_category_ids) == 1:
                where_clause += " AND rp.category_id = %s"%(partner_category_ids[0])
            else:
                where_clause += " AND rp.category_id in %s"%(tuple(partner_category_ids),)
        if self.account_type_ids:
            account_type_ids = self.account_type_ids.ids
            if len(account_type_ids) == 1:
                where_clause += " AND aa.user_type_id = %s"%(account_type_ids[0])
            else:
                where_clause += " AND aa.user_type_id in %s"%(tuple(account_type_ids),)
        if self.date_from:
            if opening:
                where_clause += " AND am.date < '%s'"%(self.date_from)
            elif detail:
                where_clause += " AND am.date >= '%s'"%(self.date_from)
        if self.date_to:
            where_clause += " AND am.date <= '%s'"%(self.date_to)
        if self.target_move == 'posted':
            where_clause += " AND am.state = 'posted'"
        return where_clause, where_params
    
    @api.model
    def view_report_details(self, option, partner_id):
        r = self.env['account.partner.ledger'].search([('id', '=', option[0])])
        print (partner_id)
        opening_where, open_param = r.get_where_condition(opening=True)
        print (open_param)
        where, param = r.get_where_condition(detail=True)
        if r.date_from:
            param.extend(open_param)
            sql = ("""WITH data AS (SELECT
                            0 AS lid,
                            aml.partner_id AS partner_id,
                            0 AS move_id, 
                            aml.account_id AS account_id,
                            '' as account_name,
                            '"""+ r.date_from.strftime('%Y-%m-%d') + """' AS ldate,
                            '' AS lcode,
                            aml.currency_id, 
                            SUM(aml.amount_currency) as amount_currency,
                            'Initial Balance' AS lref,
                            'Initial Balance' AS ref,
                            'Initial Balance' AS lname, 
                            COALESCE(sum(aml.debit),0) AS debit,
                            COALESCE(sum(aml.credit),0) AS credit, 
                            COALESCE(SUM(aml.balance),0) AS balance,
                            '' AS move_name,
                            c.symbol AS currency_code,
                            c.position AS currency_position,
                            rp.name AS partner_name
                        FROM account_move_line aml
                            LEFT JOIN account_move am ON (aml.move_id=am.id)
                            LEFT JOIN account_account aa ON (aml.account_id=aa.id)
                            LEFT JOIN res_currency c ON (aml.currency_id=c.id)
                            LEFT JOIN res_partner rp ON (aml.partner_id=rp.id)
                            LEFT JOIN account_account_type aat ON aat.id = aa.user_type_id
                        WHERE
                            %s AND aml.partner_id = %s
                        GROUP BY
                            aml.account_id,
                            aa.name,
                            aml.currency_id,
                            c.symbol,
                            c.position,
                            aml.partner_id,
                            rp.name
                        UNION
                        SELECT
                            aml.id AS lid,
                            aml.partner_id AS partner_id,
                            am.id AS move_id,
                            aml.account_id AS account_id,
                            aa.name as account_name,
                            aml.date AS ldate,
                            aj.code AS lcode,
                            aml.currency_id, 
                            aml.amount_currency,
                            aml.ref AS lref,
                            aml.ref AS ref,
                            aml.name AS lname, 
                            COALESCE(aml.debit,0) AS debit,
                            COALESCE(aml.credit,0) AS credit, 
                            COALESCE(SUM(aml.balance),0) AS balance,
                            am.name AS move_name,
                            c.symbol AS currency_code,
                            c.position AS currency_position,
                            rp.name AS partner_name
                        FROM 
                            account_move_line aml
                            LEFT JOIN account_move am ON (aml.move_id=am.id)
                            LEFT JOIN account_account aa ON (aml.account_id=aa.id)
                            LEFT JOIN res_currency c ON (aml.currency_id=c.id)
                            LEFT JOIN res_partner rp ON (aml.partner_id=rp.id)
                            LEFT JOIN account_journal aj ON (aml.journal_id=aj.id)
                            LEFT JOIN account_account_type aat ON aat.id = aa.user_type_id
                        WHERE
                            %s AND aml.partner_id = %s
                        GROUP BY
                            aml.id,
                            am.id,
                            aml.account_id,
                            aa.name,
                            aml.date,
                            aj.code,
                            aml.currency_id,
                            aml.amount_currency,
                            aml.ref,
                            aml.name,
                            am.name,
                            c.symbol,
                            c.position,
                            rp.name
                        ORDER BY
                            ldate asc, lcode)
                    SELECT
                        lid,
                        partner_id,
                        move_id,
                        account_id,
                        account_name,
                        to_char(ldate, 'DD/MM/YYYY') as ldate,
                        lcode,
                        currency_id, 
                        amount_currency,
                        lref,
                        ref,
                        lname, 
                        debit,
                        credit,
                        sum(balance) over (order by ldate asc,lcode rows between unbounded preceding and current row) as balance, 
                        move_name,
                        currency_code,
                        currency_position,
                        partner_name
                    FROM
                        data
                    """%(opening_where, partner_id, where, partner_id))
            self._cr.execute(sql, param)
        else:
            sql = ('''WITH data AS (SELECT
                            aml.id AS lid,
                            aml.partner_id AS partner_id,
                            am.id AS move_id,
                            am.ref as ref,
                            aml.account_id AS account_id,
                            aa.name as account_name,
                            aml.date AS ldate,
                            aj.code AS lcode,
                            aml.currency_id, 
                            aml.amount_currency,
                            aml.ref AS lref,
                            aml.name AS lname, 
                            COALESCE(aml.debit,0) AS debit,
                            COALESCE(aml.credit,0) AS credit, 
                            COALESCE(SUM(aml.balance),0) AS balance,
                            am.name AS move_name,
                            c.symbol AS currency_code,
                            c.position AS currency_position,
                            rp.name AS partner_name
                        FROM
                            account_move_line aml
                            LEFT JOIN account_move am ON (aml.move_id=am.id)
                            LEFT JOIN res_currency c ON (aml.currency_id=c.id)
                            LEFT JOIN res_partner rp ON (aml.partner_id=rp.id)
                            LEFT JOIN account_journal aj ON aml.journal_id=aj.id
                            LEFT JOIN account_account aa ON (aml.account_id = aa.id)
                            LEFT JOIN account_account_type aat ON aat.id = aa.user_type_id
                        WHERE
                            %s AND aml.partner_id = %s
                        GROUP BY
                            aml.id,
                            am.id,
                            aml.account_id,
                            aa.name,
                            aml.date,
                            aj.code,
                            aml.currency_id,
                            aml.amount_currency,
                            aml.ref,
                            aml.name,
                            am.name,
                            c.symbol,
                            c.position,
                            rp.name
                        ORDER BY
                            ldate asc, lcode)
                    SELECT
                        lid,
                        partner_id,
                        move_id,
                        account_id,
                        account_name,
                        to_char(ldate, 'DD/MM/YYYY') as ldate,
                        lcode,
                        currency_id, 
                        amount_currency,
                        lref,
                        ref,
                        lname, 
                        debit,
                        credit,
                        sum(balance) over (order by ldate asc,lcode rows between unbounded preceding and current row) as balance, 
                        move_name,
                        currency_code,
                        currency_position,
                        partner_name
                    FROM
                        data
                        '''%(where, partner_id))
            self._cr.execute(sql, param)
        data = self._cr.dictfetchall()
        currency = self._get_currency()
        return {
            'report_lines': data,
            'currency': currency
            }
    
    @api.model
    def view_report(self, option):
        r = self.env['account.partner.ledger'].search([('id', '=', option[0])])
        where, where_param = r.get_where_condition()
        sql = """SELECT
                    rp.id as id,
                    rp.name as name,
                    SUM(aml.debit) as debit,
                    SUM(aml.credit) as credit,
                    SUM(aml.debit - aml.credit) as balance
                FROM
                    account_move_line aml
                    LEFT JOIN account_move am ON am.id = aml.move_id
                    LEFT JOIN res_partner rp ON rp.id = aml.partner_id
                    LEFT JOIN account_account aa ON aa.id = aml.account_id
                    LEFT JOIN account_account_type aat ON aat.id = aa.user_type_id
                WHERE
                    %s
                GROUP BY
                    rp.id,
                    rp.name
        """%(where)
        self._cr.execute(sql, where_param)
        data = self._cr.dictfetchall()
        sum_sql =  """SELECT
                            SUM(aml.debit) as debit,
                            SUM(aml.credit) as credit,
                            SUM(aml.debit - aml.credit) as balance
                        FROM
                            account_move_line aml
                            LEFT JOIN account_move am ON am.id = aml.move_id
                            LEFT JOIN res_partner rp ON rp.id = aml.partner_id
                            LEFT JOIN account_account aa ON aa.id = aml.account_id
                            LEFT JOIN account_account_type aat ON aat.id = aa.user_type_id
                        WHERE
                            %s
        """%(where)
        self._cr.execute(sum_sql, where_param)
        total_data = self._cr.dictfetchall()
        filters = self.get_filter(option)
        currency = self._get_currency()
        return {
            'name': "Partner Ledger",
            'type': 'ir.actions.client',
            'tag': 'p_l',
            'filters': filters,
            'report_lines': data,
            'debit_total': total_data[0].get('debit', 0.00),
            'credit_total': total_data[0].get('credit', 0.00),
            'debit_balance': total_data[0].get('balance', 0.00),
            'currency': currency,
            'wiz_id': r.id
        }

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
            filters['accounts'] = ['All Payable and Receivable']
        if data.get('target_move'):
            filters['target_move'] = data.get('target_move').capitalize()
        if data.get('date_from'):
            filters['date_from'] = data.get('date_from')
        if data.get('date_to'):
            filters['date_to'] = data.get('date_to')
    
        filters['company_id'] = ''
        filters['accounts_list'] = data.get('accounts_list')
        filters['journals_list'] = data.get('journals_list')
    
        filters['company_name'] = data.get('company_name')
    
        if data.get('partners'):
            filters['partners'] = self.env['res.partner'].browse(
                data.get('partners')).mapped('name')
        else:
            filters['partners'] = ['All']
    
        if data.get('reconciled') == 'unreconciled':
            filters['reconciled'] = 'Unreconciled'
    
        if data.get('account_type', []):
            filters['account_type'] = self.env['account.account.type'].browse(data.get('account_type', [])).mapped('name')
        else:
            filters['account_type'] = ['Receivable and Payable']
    
        if data.get('partner_tags', []):
            filters['partner_tags'] = self.env['res.partner.category'].browse(
                data.get('partner_tags', [])).mapped('name')
        else:
            filters['partner_tags'] = ['All']
    
        filters['partners_list'] = data.get('partners_list')
        filters['category_list'] = data.get('category_list')
        filters['account_type_list'] = data.get('account_type_list')
        filters['target_move'] = data.get('target_move').capitalize()
        return filters
    
    def get_filter_data(self, option):
        r = self.env['account.partner.ledger'].search([('id', '=', option[0])])
        default_filters = {}
        company_id = self.env.company
        company_domain = [('company_id', '=', company_id.id)]
        journals = r.journal_ids if r.journal_ids else self.env['account.journal'].search(company_domain)
        accounts = self.account_ids if self.account_ids else self.env['account.account'].search(company_domain)
    
        partner = r.partner_ids if r.partner_ids else self.env[
            'res.partner'].search([])
        categories = self.partner_category_ids if self.partner_category_ids \
            else self.env['res.partner.category'].search([])
        account_types = self.env['account.account.type'].search([('type', 'in', ('receivable', 'payable'))])
    
        filter_dict = {
            'journal_ids': r.journal_ids.ids,
            'account_ids': r.account_ids.ids,
            'company_id': company_id.id,
            'date_from': r.date_from,
            'date_to': r.date_to,
            'target_move': r.target_move,
            'journals_list': [(j.id, j.name, j.code) for j in journals],
            'accounts_list': [(a.id, a.name) for a in accounts],
            'company_name': company_id and company_id.name,
            'partners': r.partner_ids.ids,
            'reconciled': r.reconciled,
            'account_type': r.account_type_ids.ids,
            'partner_tags': r.partner_category_ids.ids,
            'partners_list': [(p.id, p.name) for p in partner],
            'category_list': [(c.id, c.name) for c in categories],
            'account_type_list': [(t.id, t.name) for t in account_types],
    
        }
        filter_dict.update(default_filters)
        return filter_dict
    #
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
                          self.env.company.currency_id.position, lang,
                          self.env.company.currency_id.decimal_places]
        return currency_array

    def get_dynamic_xlsx_report(self, data, response, report_data, dfr_data):
        report_data = json.loads(report_data)
        filters = json.loads(data)
        dfr_data = json.loads(dfr_data)
        wiz_id = dfr_data['wiz_id']
        wiz_obj = self.env['account.partner.ledger'].browse(wiz_id)

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        cell_format = workbook.add_format(
            {'align': 'center', 'bold': True,
             'border': 0
             })
        sheet = workbook.add_worksheet()
        head = workbook.add_format({'align': 'center', 'bold': True,
                                    'font_size': 20})

        txt = workbook.add_format({'font_size': 10, 'border': 1})
        sub_heading_sub = workbook.add_format(
            {'align': 'center', 'bold': True, 'font_size': 10,
             'border': 1,
             'border_color': 'black'})
        sheet.merge_range('A1:H2',
                          filters.get('company_name') + ':' + 'Partner Ledger',
                          head)
        date_head = workbook.add_format({'align': 'center', 'bold': True,
                                         'font_size': 10})

        sheet.merge_range('A4:B4',
                          'Target Moves: ' + filters.get('target_move'),
                          date_head)

        sheet.merge_range('C4:D4', 'Account Type: ' + ', ' .join(
            [lt or '' for lt in
             filters['account_type']]),
                          date_head)
        sheet.merge_range('E3:F3', ' Partners: ' + ', '.join(
            [lt or '' for lt in
             filters['partners']]), date_head)
        sheet.merge_range('G3:H3', ' Partner Tags: ' + ', '.join(
            [lt or '' for lt in
             filters['partner_tags']]),
                          date_head)
        sheet.merge_range('A3:B3', ' Journals: ' + ', '.join(
            [lt or '' for lt in
             filters['journals']]),
                          date_head)
        sheet.merge_range('C3:D3', ' Accounts: ' + ', '.join(
            [lt or '' for lt in
             filters['accounts']]),
                          date_head)

        if filters.get('date_from') and filters.get('date_to'):
            sheet.merge_range('E4:F4', 'From: ' + filters.get('date_from'),
                              date_head)

            sheet.merge_range('G4:H4', 'To: ' + filters.get('date_to'),
                              date_head)
        elif filters.get('date_from'):
            sheet.merge_range('E4:F4', 'From: ' + filters.get('date_from'),
                              date_head)
        elif filters.get('date_to'):
            sheet.merge_range('E4:F4', 'To: ' + filters.get('date_to'),
                              date_head)

        sheet.merge_range('A5:E5', 'Partner', cell_format)
        sheet.write('F5', 'Debit', cell_format)
        sheet.write('G5', 'Credit', cell_format)
        sheet.write('H5', 'Balance', cell_format)

        row = 4
        col = 0

        sheet.set_column(0, 0, 15)
        sheet.set_column(1, 1, 15)
        sheet.set_column(2, 2, 25)
        sheet.set_column(3, 3, 15)
        sheet.set_column(4, 4, 36)
        sheet.set_column(5, 5, 15)
        sheet.set_column(6, 6, 15)
        sheet.set_column(7, 7, 15)

        for report in report_data:

            row += 1
            sheet.merge_range(row, col + 0, row, col + 4, report['name'],
                              sub_heading_sub)
            sheet.write(row, col + 5, report['debit'], sub_heading_sub)
            sheet.write(row, col + 6, report['credit'], sub_heading_sub)
            sheet.write(row, col + 7, report['balance'], sub_heading_sub)
            row += 1
            sheet.write(row, col + 0, 'Date', cell_format)
            sheet.write(row, col + 1, 'JRNL', cell_format)
            sheet.write(row, col + 2, 'Account', cell_format)
            sheet.write(row, col + 3, 'Move', cell_format)
            sheet.write(row, col + 4, 'Entry Label', cell_format)
            sheet.write(row, col + 5, 'Debit', cell_format)
            sheet.write(row, col + 6, 'Credit', cell_format)
            sheet.write(row, col + 7, 'Balance', cell_format)
            for r_rec in wiz_obj.view_report_details([wiz_obj.id], report['id'])['report_lines']:
                row += 1
                sheet.write(row, col + 0, r_rec['ldate'], txt)
                sheet.write(row, col + 1, r_rec['lcode'], txt)
                sheet.write(row, col + 2, r_rec['account_name'], txt)
                sheet.write(row, col + 3, r_rec['move_name'], txt)
                sheet.write(row, col + 4, r_rec['lname'], txt)
                sheet.write(row, col + 5, r_rec['debit'], txt)
                sheet.write(row, col + 6, r_rec['credit'], txt)
                sheet.write(row, col + 7, r_rec['balance'], txt)

        workbook.close()
        output.seek(0)
        response.stream.write(output.read())
        output.close()
