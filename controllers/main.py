# -*- coding: utf-8 -*-

from odoo import http
from odoo.http import content_disposition, request
import io
import xlsxwriter


class SaleExcelReportController(http.Controller):
    @http.route([
        '/sale/excel_report/<model("abc.muti.report"):wizard>',
    ], type='http', auth="user", csrf=False)
    def get_sale_excel_report(self, wizard=None, **args):


        response = request.make_response(
            None,
            headers=[
                ('Content-Type', 'application/vnd.ms-excel'),
                ('Content-Disposition', content_disposition('ABC Custom Report' + '.xlsx'))
            ]
        )

        # create workbook object from xlsxwriter library
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        # create some style to set up the font type, the font size, the border, and the aligment
        title_style = workbook.add_format({'font_name': 'Times', 'font_size': 14, 'bold': True, 'align': 'center'})
        header_style = workbook.add_format(
            {'font_name': 'Times', 'bold': True, 'left': 1, 'bottom': 1, 'right': 1, 'top': 1, 'align': 'center'})
        text_style = workbook.add_format(
            {'font_name': 'Times', 'left': 1, 'bottom': 1, 'right': 1, 'top': 1, 'align': 'left'})
        number_style = workbook.add_format(
            {'font_name': 'Times', 'left': 1, 'bottom': 1, 'right': 1, 'top': 1, 'align': 'right'})

        # loop all selected user/salesperson
        for company in wizard.company_id:
            # create worksheet/tab per salesperson
            sheet = workbook.add_worksheet(company.name)
            # set the orientation to landscape
            sheet.set_landscape()
            # set up the paper size, 9 means A4
            sheet.set_paper(9)
            # set up the margin in inch
            sheet.set_margins(0.5, 0.5, 0.5, 0.5)

            # set up the column width
            sheet.set_column('A:A', 8)
            sheet.set_column('B:P', 15)

            # the report title
            # merge the A1 to E1 cell and apply the style font size : 14, font weight : bold
            sheet.merge_range('A1:P1', 'Sales Order', title_style)

            # table title
            sheet.write(1, 0, 'Number', header_style)
            sheet.write(1, 1, 'Order Date', header_style)
            sheet.write(1, 2, 'Delivery Date', header_style)
            sheet.write(1, 3, 'Expected Date', header_style)
            sheet.write(1, 4, 'Customer', header_style)
            sheet.write(1, 5, 'Salesperson', header_style)
            sheet.write(1, 6, 'Next Activity', header_style)
            sheet.write(1, 7, 'Sales Team', header_style)
            sheet.write(1, 8, 'Warehouse', header_style)
            sheet.write(1, 9, 'Company', header_style)
            sheet.write(1, 10, 'Untaxed Amount', header_style)
            sheet.write(1, 11, 'Taxes', header_style)
            sheet.write(1, 12, 'Total', header_style)
            sheet.write(1, 13, 'Invoice Status', header_style)
            sheet.write(1, 14, 'Agents', header_style)
            sheet.write(1, 15, 'Tags', header_style)


            row = 2
            number = 1

            # search the sales order
            orders = request.env['sale.order'].search(
                [('company_id', '=', company.id), ('date_order', '>=', wizard.start_date),
                 ('date_order', '<=', wizard.end_date)])
            for order in orders:
                # the report content
                sheet.write(row, 0, number, text_style)
                sheet.write(row, 1, order.name, text_style)
                sheet.write(row, 2, str(order.date_order), text_style)
                sheet.write(row, 3, order.partner_id.name, text_style)
                sheet.write(row, 4, order.amount_total, number_style)


                row += 1
                number += 1


        # return the excel file as a response, so the browser can download it
        workbook.close()
        output.seek(0)
        response.stream.write(output.read())
        output.close()

        return response