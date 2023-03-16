# -*- coding: utf-8 -*-

from odoo import models, fields, api, exceptions, _

class SaleWizard(models.TransientModel):
    _name = 'abc.muti.report'


    company_id = fields.Many2many('res.company', string='Company')
    start_date = fields.Date("Start Date")
    end_date = fields.Date("End Date")


    def get_excel_report(self):
        # redirect to /sale/excel_report controller to generate the excel file
        return {
            'type': 'ir.actions.act_url',
            'url': '/sale/excel_report/%s' % (self.id),
            'target': 'new',
        }