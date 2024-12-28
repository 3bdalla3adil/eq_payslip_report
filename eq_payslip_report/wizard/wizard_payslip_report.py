# -*- coding: utf-8 -*-
##############################################################################
#
# Copyright 2019 EquickERP
#
##############################################################################

import xlsxwriter
import base64
from odoo import fields, models, api, _


class wizard_payslip_report(models.TransientModel):
    _name = 'wizard.payslip.report'
    _description = 'Wizard Payslip Report'

    xls_file = fields.Binary(string='Download')
    name = fields.Char(string='File name', size=64)
    state = fields.Selection([('choose', 'choose'),
                              ('download', 'download')], default="choose", string="Status")
    payslip_ids = fields.Many2many('hr.payslip', string="Payslip Ref.",
                                   default=lambda self: self._context.get('active_ids'))

    def print_instruction_xlsx(self):
        slip_ids = self.payslip_ids
        import datetime
        current_month = datetime.datetime.now().strftime('%B')
        current_year = datetime.datetime.now().year

        report_title = f"Leave & Joining Report - {current_month} {current_year}"

        xls_filename = 'Payslip Report.xlsx'
        workbook = xlsxwriter.Workbook(xls_filename)
        worksheet = workbook.add_worksheet("Batch Payslip Report")

        text_center = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        text_center.set_text_wrap()
        font_bold_center_18 = workbook.add_format(
            {'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 18})
        font_bold_center_14 = workbook.add_format(
            {'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
        font_bold_center = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
        font_bold_right = workbook.add_format({'bold': True})
        font_bold_right.set_num_format('###0.00')
        number_format = workbook.add_format()
        number_format.set_num_format('###0.00')

        worksheet.set_column('A:AZ', 18)
        worksheet.set_row(0, 40)

        # Insert company name
        worksheet.merge_range(0, 0, 0, 10, self.env.user.company_id.name, text_center)

        # Insert report title with current month and year
        worksheet.set_row(1, 40)
        worksheet.merge_range(1, 0, 1, 10, report_title, font_bold_center_18)
        worksheet.set_row(3, 30)
        worksheet.merge_range(3, 0, 3, 10, "Medical Leave & Absent", font_bold_center_14)

        row = 5
        worksheet.merge_range(row, 0, row, 0, 'Sr No', font_bold_center)
        worksheet.merge_range(row, 1, row, 1, 'Employee Name', font_bold_center)
        worksheet.merge_range(row, 2, row, 2, 'No of Days', font_bold_center)
        worksheet.merge_range(row, 3, row, 3, 'Date', font_bold_center)
        worksheet.merge_range(row, 4, row, 4, 'Remarks', font_bold_center)

        row += 2  # Two blank rows for spacing
        worksheet.set_row(row, 30)
        worksheet.set_row(row + 1, 30)

        row += 2  # Start data rows after blank rows
        worksheet.merge_range(row, 0, row, 10, "Workers vacation/ Resigned /Joined and Returned", font_bold_center_14)
        row += 1
        worksheet.write(row, 0, 'Sr No', font_bold_center)
        worksheet.write(row, 1, 'Employee Name', font_bold_center)
        worksheet.write(row, 2, 'Joining Date', font_bold_center)
        worksheet.write(row, 3, 'Last Working Date', font_bold_center)
        worksheet.write(row, 4, 'Remarks', font_bold_center)

        row += 3  # Two blank rows for spacing
        worksheet.set_row(row, 30)
        worksheet.set_row(row + 1, 30)

        row += 2  # Start data rows after blank rows
        worksheet.merge_range(row, 0, row, 10, "Ticket Booking", font_bold_center_14)
        row += 1
        worksheet.write(row, 0, 'Sr No', font_bold_center)
        worksheet.write(row, 1, 'Employee Name', font_bold_center)
        worksheet.write(row, 2, 'Travel From', font_bold_center)
        worksheet.write(row, 3, 'Travel To', font_bold_center)
        worksheet.write(row, 4, 'Remarks', font_bold_center)

        row += 3  # Two blank rows for spacing
        worksheet.set_row(row, 30)
        worksheet.set_row(row + 1, 30)

        row += 2  # Start data rows after blank rows
        worksheet.merge_range(row, 0, row, 10, "Other Notes", font_bold_center_14)
        row += 1
        worksheet.write(row, 0, 'Sr No', font_bold_center)
        worksheet.write(row, 1, 'Employee Name', font_bold_center)
        worksheet.write(row, 2, '', font_bold_center)  # Blank column
        worksheet.write(row, 3, 'Amount', font_bold_center)
        worksheet.write(row, 4, 'Remarks', font_bold_center)

        row += 3  # Two blank rows for spacing
        worksheet.set_row(row, 30)
        worksheet.set_row(row + 1, 30)

        row += 2  # Final rows
        worksheet.write(row, 1, 'Prepared by', font_bold_center)
        worksheet.write(row, 4, 'Verified by', font_bold_center)

        workbook.close()
        action = self.env.ref('eq_payslip_report.action_wizard_payslip_report').read()[0]
        action['res_id'] = self.id
        self.write({'state': 'download',
                    'name': xls_filename,
                    'xls_file': base64.b64encode(open(xls_filename, 'rb').read())})
        return action

        def print_report_xls(self):
        slip_ids = self.payslip_ids
        xls_filename = 'Payslip Report.xlsx'
        workbook = xlsxwriter.Workbook(xls_filename)
        worksheet = workbook.add_worksheet("Batch Payslip Report")

        # Define formats
        text_center = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        text_center.set_text_wrap()
        font_bold_center = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
        font_bold_right = workbook.add_format({'bold': True})
        font_bold_right.set_num_format('###0.00')
        number_format = workbook.add_format()
        number_format.set_num_format('###0.00')
        category_format = workbook.add_format(
            {'bold': True, 'font_size': 16, 'font_color': '#D2691E'})  # Chocolate accent
        total_format = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'align': 'right'})

        # Job categories
        job_categories = {
            "Company Staff": {
                "Office Staff": ["Admin", "Accountant", "Store Keeper", "IT Supervisor"],
                "Engineers": ["Engineer", "Supervisor"],
                "PRO/Drivers": ["Mandoop"],
                "TeaBoy": ["House Boy"],
            },
            "Company Labour": {
                "Technicians": ["AC Technician", "Mason", "Plumber", "Electrician", "Helper", "General Technician"],
                "Gardener": ["Gardner"],
                "Watchmen": ["Watchman"],
                "Drivers": ["Driver"],
            },
            "Mama Villa Labour": {
                "Mama Villa": ["Mama Villa Labour", "Mama Villa Gardner", "Mama Villa Cook", "Mama Villa Driver"]
            },
            "Dr. Badria": {
                "Dr. Badria": ["Dr.Badria Cook", "Dr.Badria Housemaid"]
            },
        }

        # Column settings and header
        worksheet.set_column('A:AZ', 18)
        worksheet.set_row(0, 18)
        worksheet.write(0, 1, 'Company Name', font_bold_center)
        worksheet.merge_range(0, 2, 0, 3, self.env.user.company_id.name or '', text_center)

        # Header row
        row = 3
        worksheet.merge_range(row, 0, row + 1, 0, 'No #', font_bold_center)
        worksheet.merge_range(row, 1, row + 1, 1, 'Payslip Ref', font_bold_center)
        worksheet.merge_range(row, 2, row + 1, 2, 'Employee', font_bold_center)
        worksheet.merge_range(row, 3, row + 1, 3, 'Designation', font_bold_center)
        col = 4
        worksheet.set_row(row + 1, 30)
        result = self.get_header(slip_ids)
        col_lst = []

        # Create dynamic headers
        for item in result:
            for categ_id, salary_rule_ids in item.items():
                if not salary_rule_ids:
                    continue
                if len(salary_rule_ids) == 1:
                    worksheet.write(row, col, categ_id.name, font_bold_center)
                    worksheet.write(row + 1, col, salary_rule_ids[0].name, text_center)
                    col += 1
                    col_lst.append(salary_rule_ids[0])
                else:
                    rule_count = len(salary_rule_ids) - 1
                    worksheet.merge_range(row, col, row, col + rule_count, categ_id.name, font_bold_center)
                    for rule_id in salary_rule_ids.sorted(key=lambda l: l.sequence):
                        worksheet.write(row + 1, col, rule_id.name, text_center)
                        col += 1
                        col_lst.append(rule_id)
        worksheet.merge_range(row, col, row + 1, col, 'Payment Method', font_bold_center)
        col_lst.append('payment_method')
        row += 3

        # Group and write data
        sr_no = 1
        grand_totals = [0] * len(col_lst)
        for main_category, subcategories in job_categories.items():
            worksheet.write(row, 0, main_category, category_format)
            row += 1
            for subcategory, job_titles in subcategories.items():
                worksheet.write(row, 0, subcategory, font_bold_center)
                row += 1
                subcategory_totals = [0] * len(col_lst)
                for payslip in slip_ids.filtered(lambda p: p.employee_id.job_id.name in job_titles):
                    worksheet.write(row, 0, sr_no, text_center)
                    worksheet.write(row, 1, payslip.number)
                    worksheet.write(row, 2, payslip.employee_id.name)
                    worksheet.write(row, 3, payslip.employee_id.job_id.name or '')
                    col = 4
                    for col_index, col_rule_id in enumerate(col_lst):
                        if col_rule_id == 'payment_method':
                            worksheet.write(row, col, payslip.contract_id.payment_method or '')
                        else:
                            line_id = payslip.line_ids.filtered(lambda l: l.salary_rule_id.id == col_rule_id.id)
                            amount = line_id.total or 0.0
                            worksheet.write(row, col, amount, number_format)
                            subcategory_totals[col_index] += amount
                            grand_totals[col_index] += amount
                        col += 1
                    row += 1
                    sr_no += 1

                # Write subcategory total
                worksheet.write(row, 3, "Total", total_format)
                col = 4
                for total in subcategory_totals:
                    worksheet.write(row, col, total, number_format)
                    col += 1
                row += 1

        # Write grand total
        worksheet.write(row, 3, "Grand Total", total_format)
        col = 4
        for total in grand_totals:
            worksheet.write(row, col, total, number_format)
            col += 1

        workbook.close()
        action = self.env.ref('eq_payslip_report.action_wizard_payslip_report').read()[0]
        action['res_id'] = self.id
        self.write({'state': 'download',
                    'name': xls_filename,
                    'xls_file': base64.b64encode(open(xls_filename, 'rb').read())})
        return action

    #ToDo: NEED TO FIX PDF TEMPLATE
    # def print_report_pdf(self):
    def print_report_pdf(self):
        # pass


        data = self.read()[0]
        return self.env.ref('eq_payslip_report.action_print_payslip').report_action([], data=data)

    def get_header(self, slip_ids):
        category_list_ids = self.env['hr.salary.rule.category'].search([])
        # find all the rule by category
        col_by_category = {}
        for payslip in slip_ids:
            for line in payslip.line_ids:
                col_by_category.setdefault(line.category_id, [])
                col_by_category[line.category_id] += line.salary_rule_id.ids
        for categ_id, rule_ids in col_by_category.items():
            col_by_category[categ_id] = self.env['hr.salary.rule'].browse(set(rule_ids))
        # make the category wise rule
        result = []
        for categ_id in category_list_ids:
            rule_ids = col_by_category.get(categ_id)
            if not rule_ids:
                continue
            result.append({categ_id: rule_ids.sorted(lambda l: l.sequence)})
        return result

    def action_go_back(self):
        action = self.env.ref('eq_payslip_report.action_wizard_payslip_report').read()[0]
        action['res_id'] = self.id
        self.write({'state': 'choose'})
        return action


class eq_payslip_report_report_payslip_template(models.AbstractModel):
    _name = 'report.eq_payslip_report.report_payslip_template'

    @api.model
    def _get_report_values(self, docids, data=None):
        report = self.env['ir.actions.report']._get_report_from_name('eq_payslip_report.report_payslip_template')
        wizard_id = self.env['wizard.payslip.report'].browse(data.get('id'))
        slip_ids = wizard_id.payslip_ids
        get_header = wizard_id.get_header(slip_ids)
        get_rule_list = []
        for header_dict in get_header:
            for rule_ids in header_dict.values():
                for rule_id in rule_ids:
                    get_rule_list.append(rule_id)
        return {
            'doc_ids': self.ids,
            'doc_model': report,
            'docs': slip_ids,
            'data': data,
            'slip_ids': slip_ids,
            '_get_header': get_header,
            '_get_rule_list': get_rule_list
        }

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
