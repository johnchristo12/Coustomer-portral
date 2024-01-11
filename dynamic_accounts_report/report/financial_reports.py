# -*- coding: utf-8 -*-
from odoo import api, models, _


class InsReportBalanceSheet(models.AbstractModel):
    _name = 'report.dynamic_accounts_report.balance_sheet'

    @api.model
    def _get_report_values(self, docids, data=None):
        if self.env.context.get('bs_report'):
            if data.get('report_data'):
                data.update({
                    'Filters': data.get('report_data')['filters'],
                    'report_lines': data.get('report_data')['bs_lines'],
                    'report_name': data.get('report_name'),
                    'title': data.get('report_data')['name'],
                    'company': self.env.company,
                    'has_comp': data.get('report_data')['has_comp']
                })
        return data
