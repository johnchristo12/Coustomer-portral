from odoo import api, models, _


class TrialBalance(models.AbstractModel):
    _name = 'report.dynamic_accounts_report.trial_balance'

    @api.model
    def _get_report_values(self, docids, data=None):
        if self.env.context.get('trial_pdf_report'):

            if data.get('report_data'):
                data.update({'account_data': data.get('report_data')['report_lines'],
                             'Filters': data.get('report_data')['filters'],
                             'debit_total': data.get('report_data')['debit_total'],
                             'credit_total': data.get('report_data')['credit_total'],
                             'op_debit_total': data.get('report_data')['op_debit_total'],
                             'op_credit_total': data.get('report_data')['op_credit_total'],
                             'cl_debit_total': data.get('report_data')['cl_debit_total'],
                             'cl_credit_total': data.get('report_data')['cl_credit_total'],
                             'company': self.env.company,
                             })
        return data
