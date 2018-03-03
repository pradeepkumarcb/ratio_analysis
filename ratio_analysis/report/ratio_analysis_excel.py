# -*- coding: utf-8 -*-
##############################################################################
#
#    ODOO, Open Source Management Solution
#    Copyright (C) 2016-TODAY Steigend IT Solutions
#    For more details, check COPYRIGHT and LICENSE files
#
##############################################################################
from datetime import datetime
from odoo import api, models, _
from odoo.addons.report_xlsx.report.report_xlsx import ReportXlsx
from dateutil.relativedelta import relativedelta

class RatioReportXlsx(ReportXlsx):
   
    def get_current_month_movelines(self,date_from, date_to):
        current_result = []
        config_records = self.env['profitability.report.configruration'].search([])
        direct_income_accounts = []
        indirect_income_accounts = []
        direct_expense_accounts = []
        indirect_expense_accounts = []
        sale_account_ids = []
        purchase_account_ids = []
        current_asset_accounts = []
        current_liability_accounts = []
        totalliability_accounts = []
        total_asset_accounts = []
        equity_accounts = []
        receivable_accounts = []
        payable_accounts = []
        for config in config_records:
            direct_income_accounts = config.income_account_ids
            indirect_income_accounts = config.indirect_income_account_ids
            direct_expense_accounts = config.direct_expenses_account_ids
            indirect_expense_accounts = config.indirect_expenses_account_ids
            sale_account_ids = config.sales_income_account_ids
            purchase_account_ids = config.purchase_expense_account_ids
            current_asset_types = config.current_asset_ids
            current_liability_types = config.current_liability_ids
            total_liability_types = config.totalliability_ids
            total_asset_types = config.total_asset_ids
            equity_types = config.equity_ids
            receivable_types = config.receivable_ids
            payable_types = config.payable_ids
            if current_asset_types:
                current_asset_accounts = self.env['account.account'].search([('user_type_id', 'in', current_asset_types.ids)])
            if current_liability_types:
                current_liability_accounts = self.env['account.account'].search([('user_type_id', 'in', current_liability_types.ids)])
            if total_liability_types:
                totalliability_accounts = self.env['account.account'].search([('user_type_id', 'in', total_liability_types.ids)])
            if total_asset_types:
                total_asset_accounts = self.env['account.account'].search([('user_type_id', 'in', total_asset_types.ids)])
            if equity_types:
                equity_accounts = self.env['account.account'].search([('user_type_id', 'in', equity_types.ids)])
            if receivable_types:
                receivable_accounts = self.env['account.account'].search([('user_type_id', 'in', receivable_types.ids)])
            if payable_types:
                payable_accounts = self.env['account.account'].search([('user_type_id', 'in', payable_types.ids)])
        invoices = self.env['account.invoice'].search([('type', '=', 'out_invoice'), 
                                                       ('state', '=', 'paid'),
                                                       ('date', '>=', date_from),
                                                       ('date', '<=', date_to)])
        sales_amount = 0
        gross_profit_margin = 0
        net_profit_percentage = 0
        #Sales amount
        if sale_account_ids:
            sales_income_credit_movelines = self.get_movelines(sale_account_ids, 'credit', date_from, date_to)
            sales_income_debit_movelines = self.get_movelines(sale_account_ids, 'debit', date_from, date_to)
            sales_amount = self.get_balance(sales_income_debit_movelines, sales_income_credit_movelines, 'minus')
        else:
            for invoice in invoices:
                sales_amount += invoice.amount_total
        
        #Direct income
        direct_income_credit_movelines = self.get_movelines(direct_income_accounts, 'credit', date_from, date_to)
        direct_income_debit_movelines = self.get_movelines(direct_income_accounts, 'debit', date_from, date_to)
        direct_income = 0
        direct_income = self.get_balance(direct_income_debit_movelines, direct_income_credit_movelines, 'minus')
        
        
        direct_expense = 0
        gross_profit = 0
        
        #Direct Expense
        direct_expense_credit_movelines = self.get_movelines(direct_expense_accounts, 'credit', date_from, date_to)
        direct_expense_debit_movelines = self.get_movelines(direct_expense_accounts, 'debit', date_from, date_to)
        direct_expense = self.get_balance(direct_expense_debit_movelines, direct_expense_credit_movelines, 'plus')
        gross_profit = direct_income - direct_expense
        indirect_income_credit_movelines = self.get_movelines(indirect_income_accounts, 'credit', date_from, date_to)
        indirect_income_debit_movelines = self.get_movelines(indirect_income_accounts, 'debit', date_from, date_to)
        indirect_income = 0
        indirect_income = self.get_balance(indirect_income_debit_movelines, indirect_income_credit_movelines, 'minus')
        indirect_expense = 0
        net_profit = 0
        
        indirect_expense_credit_movelines = self.get_movelines(indirect_expense_accounts, 'credit', date_from, date_to)
        indirect_expense_debit_movelines = self.get_movelines(indirect_expense_accounts, 'debit', date_from, date_to)
        indirect_expense = self.get_balance(indirect_expense_debit_movelines, indirect_expense_credit_movelines, 'plus')
        indirect_profit = indirect_income - indirect_expense
        net_profit = gross_profit + indirect_profit

        current_ratio = 0
        current_asset = 0
        current_asset_credit_movelines = self.get_movelines(current_asset_accounts, 'credit', date_from, date_to)
        current_asset_debit_movelines = self.get_movelines(current_asset_accounts, 'debit', date_from, date_to)
        current_asset = self.get_balance(current_asset_debit_movelines, current_asset_credit_movelines, 'plus')
        
        current_liability = 0
        current_liability_credit_movelines = self.get_movelines(current_liability_accounts, 'credit', date_from, date_to)
        current_liability_debit_movelines = self.get_movelines(current_liability_accounts, 'debit', date_from, date_to)
        current_liability = self.get_balance(current_liability_debit_movelines, current_liability_credit_movelines, 'plus')
        if current_liability:
            current_ratio = current_asset/current_liability * 100
        
        asset_lia_diff = 0
        net_working_capital = 0
        asset_lia_diff = current_asset - current_liability
        if sales_amount:
            gp_by_sa = gross_profit/sales_amount
            gross_profit_margin = round(gp_by_sa * 100, 2)
            np_by_sa = net_profit/sales_amount
            net_profit_percentage = round(np_by_sa * 100, 2)
            asset_lia_diff_by_sa = asset_lia_diff / sales_amount
            net_working_capital = round(asset_lia_diff_by_sa * 100, 2)
        
        total_liability = 0
        debit_equity_ratio = 0
        equity_asset_ratio = 0
        shareholder_equity = 0
        equity = 0 
        total_liability_credit_movelines = self.get_movelines(totalliability_accounts, 'credit', date_from, date_to)
        
        total_liability_debit_movelines = self.get_movelines(totalliability_accounts, 'debit', date_from, date_to)
       
        total_liability = self.get_balance(total_liability_debit_movelines, total_liability_credit_movelines, 'plus')
        
        equity_credit_movelines = self.get_movelines(equity_accounts, 'credit', date_from, date_to)
        equity_debit_movelines = self.get_movelines(equity_accounts, 'debit', date_from, date_to)
        
        equity = self.get_balance(equity_debit_movelines, equity_credit_movelines, 'plus')
        if equity:
            total_lia_by_equity = total_liability / equity
            debit_equity_ratio = round(total_lia_by_equity * 100, 2)

        total_asset_credit_movelines = self.get_movelines(total_asset_accounts, 'credit', date_from, date_to)
        total_asset_debit_movelines = self.get_movelines(total_asset_accounts, 'debit', date_from, date_to)
        
        total_asset = self.get_balance(total_asset_debit_movelines, total_asset_credit_movelines, 'plus')
        if total_asset:
            equity_by_total_asset = equity / total_asset
            equity_asset_ratio = round(equity_by_total_asset * 100, 2)
        shareholder_equity = round(total_asset - total_liability, 2)

        opening_credit_recievable_movelines = self.env['account.move.line'].search([('account_id.user_type_id', 'in', receivable_accounts.ids),
                                                                        ('debit', '=', 0),
                                                                        ('date', '<', date_from)])
        opening_debit_recievable_movelines = self.env['account.move.line'].search([('account_id.user_type_id', 'in', receivable_accounts.ids),
                                                                        ('credit', '=', 0),
                                                                        ('date', '<', date_from)])
        opening_receivable = self.get_balance(opening_debit_recievable_movelines, opening_credit_recievable_movelines, 'plus')
        closing_recievable_credit_movelines = self.get_movelines(receivable_accounts, 'credit', date_from, date_to)
        closing_recievable_debit_movelines = self.get_movelines(receivable_accounts, 'debit', date_from, date_to)
        closing_receivable = self.get_balance(closing_recievable_debit_movelines, closing_recievable_credit_movelines, 'plus')
        actual_closing_recievable = opening_receivable + closing_receivable
        average_receivable = round(opening_receivable + actual_closing_recievable / 2, 2)
        receivable_turnover_in_day = 0
        receivable_turnover = 0
        if average_receivable:
            receivable_turnover = round(sales_amount / average_receivable, 2)
        if receivable_turnover:
            receivable_turnover_in_day = round(30 / receivable_turnover, 2)
        opening_credit_payable_movelines = self.env['account.move.line'].search([('account_id.user_type_id', 'in', payable_accounts.ids),
                                                                        ('debit', '=', 0),
                                                                        ('date', '<', date_from)])
        opening_debit_payable_movelines = self.env['account.move.line'].search([('account_id.user_type_id', 'in', payable_accounts.ids),
                                                                        ('credit', '=', 0),
                                                                        ('date', '<', date_from)])
        opening_payable = self.get_balance(opening_debit_payable_movelines, opening_credit_payable_movelines, 'plus')
        closing_payable_credit_movelines = self.get_movelines(payable_accounts, 'credit', date_from, date_to)
        closing_payable_debit_movelines = self.get_movelines(payable_accounts, 'debit', date_from, date_to)
        closing_payable = self.get_balance(closing_payable_debit_movelines, closing_payable_credit_movelines, 'plus')
        
        actual_closing_payable = opening_payable + closing_payable
        average_payable = round(opening_payable + actual_closing_payable / 2, 2)
        
        
        supplier_invoices = self.env['account.invoice'].search([('type', '=', 'in_invoice'),
                                                                ('state', '=', 'paid'),
#                                                                 ('state', '=', 'open'),
                                                                ('date', '>=', date_from),
                                                                ('date', '<=', date_to)])
        total_supplier_purchase = 0
        payable_turnover_days = 0
        payable_turnover = 0
        if purchase_account_ids:
            purchase_expense_credit_movelines = self.get_movelines(purchase_account_ids, 'credit', date_from, date_to)
            purchase_expense_debit_movelines = self.get_movelines(purchase_account_ids, 'debit', date_from, date_to)
            total_supplier_purchase = self.get_balance(purchase_expense_credit_movelines, purchase_expense_debit_movelines, 'plus')
        else:
            for invoice in supplier_invoices:
                total_supplier_purchase += invoice.amount_total
        if average_payable:
            payable_turnover = round(total_supplier_purchase / average_payable, 2)
        if payable_turnover:
            payable_turnover_days = round(30 / payable_turnover, 2)
            
        operating_cycle = receivable_turnover_in_day + payable_turnover_days
        current_result.append({'gross_profit_margin': gross_profit_margin, 'net_profit_percentage':net_profit_percentage, 'current_ratio':current_ratio, 'net_working_capital':net_working_capital,\
             'debit_equity_ratio':debit_equity_ratio, 'equity_asset_ratio':equity_asset_ratio, 'shareholder_equity':shareholder_equity, 'receivable_turnover':receivable_turnover,\
             'receivable_turnover_in_day':receivable_turnover_in_day, 'average_receivable':average_receivable, 'average_payable':average_payable, 'payable_turnover':payable_turnover, \
             'payable_turnover_days':payable_turnover_days, 'operating_cycle':operating_cycle})
            
        return current_result
    
    def get_movelines(self, account_ids, flag, date_from, date_to):
    
        if account_ids:
            if flag == 'credit':
                movelines = self.env['account.move.line'].search([('account_id', 'in', account_ids.ids),
                                                                            ('debit', '=', 0),
                                                                            ('date', '>=', date_from),
                                                                            ('date', '<=', date_to)])
            else:
                movelines = self.env['account.move.line'].search([('account_id', 'in', account_ids.ids),
                                                                            ('credit', '=', 0),
                                                                            ('date', '>=', date_from),
                                                                            ('date', '<=', date_to)])
            return movelines
        else:
            return []
    
    def get_balance(self, debit_movelines, credit_movelines, sign):
        balance = 0
        total_credit = 0
        total_debit = 0
        for debit_line in debit_movelines:
            total_debit += debit_line.debit
        for credit_line in credit_movelines:
            total_credit += credit_line.credit
        print total_credit, total_debit, sign
        balance = total_debit - total_credit
        if sign == 'minus':
            balance  = balance * -1
        return balance
    
    
    
    def get_previous_month_movelines(self, date_from, date_to):
#         date_from = datetime.strptime(date_to, '%Y-%m-%d').strftime('%Y-%m-01')
#         date_from = date_to.strftime('%Y-%m-01')
        previous_result = []
        account_records = self.env['profitability.report.configruration'].search([])
        for account in account_records:
            direct_income_accounts = account.income_account_ids
            indirect_income_accounts = account.indirect_income_account_ids
            direct_expense_accounts = account.direct_expenses_account_ids
            indirect_expense_accounts = account.indirect_expenses_account_ids
            current_asset_accounts = account.current_asset_ids
            current_liability_accounts = account.current_liability_ids
            totalliability_accounts = account.totalliability_ids
            total_asset_accounts = account.total_asset_ids
            equity_accounts = account.equity_ids
            receivable_accounts = account.receivable_ids
            payable_accounts = account.payable_ids
            
#             
        invoices = self.env['account.invoice'].search([('type', '=', 'out_invoice'), 
                                                       ('state', '=', 'paid'),
                                                       ('date', '>=', date_from),
                                                       ('date', '<=', date_to)])
        sales_amount = 0
        gross_profit_margin = 0
        net_profit_percentage = 0
        for invoice in invoices:
            sales_amount += invoice.amount_total
        direct_income_credit_movelines = self.get_movelines(direct_income_accounts, 'credit', date_from, date_to)
        direct_income_debit_movelines = self.get_movelines(direct_income_accounts, 'debit', date_from, date_to)
        
        direct_income = 0
        direct_income = self.get_balance(direct_income_debit_movelines, direct_income_credit_movelines, 'minus')
        
        
        direct_expense = 0
        gross_profit = 0
        
        direct_expense_credit_movelines = self.get_movelines(direct_expense_accounts, 'credit', date_from, date_to)
        direct_expense_debit_movelines = self.get_movelines(direct_expense_accounts, 'debit', date_from, date_to)
        direct_expense = self.get_balance(direct_expense_debit_movelines, direct_expense_credit_movelines, 'plus')
        gross_profit = direct_income - direct_expense
        
        indirect_income_credit_movelines = self.get_movelines(indirect_income_accounts, 'credit', date_from, date_to)
        indirect_income_debit_movelines = self.get_movelines(indirect_income_accounts, 'debit', date_from, date_to)
        indirect_income = 0
        indirect_income = self.get_balance(indirect_income_debit_movelines, indirect_income_credit_movelines, 'minus')
        indirect_expense = 0
        net_profit = 0
        
        indirect_expense_credit_movelines = self.get_movelines(indirect_expense_accounts, 'credit', date_from, date_to)
        indirect_expense_debit_movelines = self.get_movelines(indirect_expense_accounts, 'debit', date_from, date_to)
        indirect_expense = self.get_balance(indirect_expense_debit_movelines, indirect_expense_credit_movelines, 'plus')

        indirect_profit = indirect_income - indirect_expense
        net_profit = gross_profit + indirect_profit
        
        current_ratio = 0
        current_asset = 0
        current_asset_credit_movelines = self.get_movelines(current_asset_accounts, 'credit', date_from, date_to)
        current_asset_debit_movelines = self.get_movelines(current_asset_accounts, 'debit', date_from, date_to)
        current_asset = self.get_balance(current_asset_debit_movelines, current_asset_credit_movelines, 'plus')
        
        current_liability = 0
        current_liability_credit_movelines = self.get_movelines(current_liability_accounts, 'credit', date_from, date_to)
        current_liability_debit_movelines = self.get_movelines(current_liability_accounts, 'debit', date_from, date_to)
            
        current_liability = self.get_balance(current_liability_debit_movelines, current_liability_credit_movelines, 'plus')
        if current_liability:
            current_ratio = round(current_asset/current_liability * 100, 2)
        
        asset_lia_diff = 0
        net_working_capital = 0
        asset_lia_diff = current_asset - current_liability
        if sales_amount:
            gross_profit_margin = round(gross_profit/sales_amount * 100, 2)
            net_profit_percentage = round(net_profit/sales_amount * 100, 2)
            net_working_capital = round(asset_lia_diff / sales_amount * 100, 2)
        
        total_liability = 0
        debit_equity_ratio = 0
        equity_asset_ratio = 0
        shareholder_equity = 0
        equity = 0 
        total_liability_credit_movelines = self.get_movelines(totalliability_accounts, 'credit', date_from, date_to)
        
        total_liability_debit_movelines = self.get_movelines(totalliability_accounts, 'debit', date_from, date_to)
       
        total_liability = self.get_balance(total_liability_debit_movelines, total_liability_credit_movelines, 'plus')
        
        equity_credit_movelines = self.get_movelines(equity_accounts, 'credit', date_from, date_to)
        equity_debit_movelines = self.get_movelines(equity_accounts, 'debit', date_from, date_to)
        
        equity = self.get_balance(equity_debit_movelines, equity_credit_movelines, 'plus')
        if equity:
            debit_equity_ratio = round(total_liability / equity * 100, 2)

        total_asset_credit_movelines = self.get_movelines(total_asset_accounts, 'credit', date_from, date_to)
        total_asset_debit_movelines = self.get_movelines(total_asset_accounts, 'debit', date_from, date_to)
        
        total_asset = self.get_balance(total_asset_debit_movelines, total_asset_credit_movelines, 'plus')
        if total_asset:
            equity_by_total_asset = equity / total_asset
            equity_asset_ratio = round(equity_by_total_asset * 100, 2)
        shareholder_equity = round(total_asset - total_liability, 2)

        opening_credit_recievable_movelines = self.env['account.move.line'].search([('account_id.user_type_id', 'in', receivable_accounts.ids),
                                                                        ('debit', '=', 0),
                                                                        ('date', '<', date_from)])
        opening_debit_recievable_movelines = self.env['account.move.line'].search([('account_id.user_type_id', 'in', receivable_accounts.ids),
                                                                        ('credit', '=', 0),
                                                                        ('date', '<', date_from)])
        opening_receivable = self.get_balance(opening_debit_recievable_movelines, opening_credit_recievable_movelines, 'plus')
        
        closing_recievable_credit_movelines = self.get_movelines(receivable_accounts, 'credit', date_from, date_to)
        closing_recievable_debit_movelines = self.get_movelines(receivable_accounts, 'debit', date_from, date_to)
        closing_receivable = self.get_balance(closing_recievable_debit_movelines, closing_recievable_credit_movelines)
        
        actual_closing_recievable = opening_receivable + closing_receivable
        average_receivable = opening_receivable + actual_closing_recievable / 2
        receivable_turnover = 0
        receivable_turnover_in_day = 0
        if average_receivable:
            receivable_turnover = round(sales_amount / average_receivable, 2)
        if receivable_turnover:
            receivable_turnover_in_day = round(30 / receivable_turnover, 2)

        opening_credit_payable_movelines = self.env['account.move.line'].search([('account_id.user_type_id', 'in', payable_accounts.ids),
                                                                        ('debit', '=', 0),
                                                                        ('date', '<', date_from)])
        opening_debit_payable_movelines = self.env['account.move.line'].search([('account_id.user_type_id', 'in', payable_accounts.ids),
                                                                        ('credit', '=', 0),
                                                                        ('date', '<', date_from)])
        opening_payable = self.get_balance(opening_debit_payable_movelines, opening_credit_payable_movelines, 'plus')
        closing_payable_credit_movelines = self.get_movelines(receivable_accounts, 'credit', date_from, date_to)
        closing_payable_debit_movelines = self.get_movelines(receivable_accounts, 'debit', date_from, date_to)
        closing_payable = self.get_balance(closing_payable_debit_movelines, closing_payable_credit_movelines, 'plus')
        
        actual_closing_payable = opening_payable + closing_payable
        average_payable = opening_payable + actual_closing_payable / 2
        
        
        supplier_invoices = self.env['account.invoice'].search([('type', '=', 'in_invoice'),
                                                                ('state', '=', 'paid'),
#                                                                 ('state', '=', 'open'),
                                                                ('date', '>=', date_from),
                                                                ('date', '<', date_to)])
        total_supplier_purchase = 0
        payable_turnover = 0
        payable_turnover_days = 0
        for invoice in supplier_invoices:
            total_supplier_purchase += invoice.amount_total
        if average_payable:
            payable_turnover = round(total_supplier_purchase / average_payable, 2)
        if payable_turnover:
            payable_turnover_days = round(30 / payable_turnover, 2)
            
        operating_cycle = receivable_turnover_in_day + payable_turnover_days
        

        previous_result.append({'gross_profit_margin': gross_profit_margin, 'net_profit_percentage':net_profit_percentage, 'current_ratio':current_ratio, 'net_working_capital':net_working_capital,\
             'debit_equity_ratio':debit_equity_ratio, 'equity_asset_ratio':equity_asset_ratio, 'shareholder_equity':shareholder_equity, 'receivable_turnover':receivable_turnover,\
             'receivable_turnover_in_day':receivable_turnover_in_day, 'average_receivable':average_receivable, 'average_payable':average_payable, 'payable_turnover':payable_turnover, \
             'payable_turnover_days':payable_turnover_days, 'operating_cycle':operating_cycle})
        
        return previous_result
        
        
    
    
    def generate_xlsx_report(self, workbook, data, lines):
        sheet = workbook.add_worksheet("Ratio Analysis Report")
        sheet.set_column(1, 15, 15)
        format1 = workbook.add_format({'font_size': 14, 'bottom': True, 'right': True, 'left': True, 'top': True, 'align': 'center', 'bold': True, })
        format11 = workbook.add_format({'font_size': 12, 'align': 'center', 'right': True, 'left': True, 'bottom': True, 'top': True, 'bold': True})
        format21 = workbook.add_format({'font_size': 10, 'align': 'center', 'right': True, 'left': True,'bottom': True, 'top': True, 'bold': True})
        format3 = workbook.add_format({'bottom': True, 'top': True, 'font_size': 10, 'num_format': '#,###'})
        font_size_8 = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 8})
        red_mark = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 8, 'bg_color': 'red'})         
        justify = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 12})
        format13 = workbook.add_format({'num_format': '#,##0.00','font_size': 12, 'align': 'right', 'right': True, 'left': True, 'bottom': True, 'top': True})
        format12 = workbook.add_format({'font_size': 12, 'align': 'right', 'right': True, 'left': True, 'bottom': True, 'top': True})
        format14 = workbook.add_format({'font_size': 12, 'align': 'left', 'right': True, 'left': True, 'bottom': True, 'top': True})

        format3.set_align('center')
        font_size_8.set_align('center')
        justify.set_align('justify')
        format1.set_align('center')
        red_mark.set_align('center')

        date_from = data['form'].get('from_date')
        date_to = data['form'].get('to_date')
        convert_date = datetime.strptime(date_to, '%Y-%m-%d')
        
        current_result = self.get_current_month_movelines(date_from, date_to)
        
        sheet.merge_range('B2:E2', 'RATIO ANALYSIS', format11)
        sheet.merge_range('B3:E3', 'FOR THE MONTHS OF ' + convert_date.strftime("%b") + '-' + str(convert_date.year), format1)
        
        
        sheet.merge_range('A1:A3', ' ', format11)
        sheet.merge_range('A4:C4', '', format11)
        sheet.merge_range('D4:E4', '', format11)
        sheet.merge_range('A5:C5', 'RATIO OR OTHER MEASUREMENT', format11)
        sheet.merge_range(5, 0, 5, 2, 'Gross Profit Margin (%)', format14)
        sheet.merge_range(6, 0, 6, 2, 'Net Profit (%)', format14)
        sheet.merge_range(7, 0, 7, 2, 'Current Ratio', format14)
        sheet.merge_range(8, 0, 8, 2, 'Net Working Capital', format14)
        sheet.merge_range(9, 0, 9, 2, 'Debt Equity Ratio', format14)
        sheet.merge_range(10, 0, 10, 2, 'Equity-Assets Ratio', format14)
        sheet.merge_range(11, 0, 11, 2, 'Shareholder Equity ', format14)
        sheet.merge_range(12, 0, 12, 2, 'Receivable Turnover', format14)
        sheet.merge_range(13, 0, 13, 2, 'Receivable Turnover in day', format14)
        sheet.merge_range(14, 0, 14, 2, 'Avg Receivable', format14)
        sheet.merge_range(15, 0, 15, 2, 'Payable Turnover', format14)
        sheet.merge_range(16, 0, 16, 2, 'Payable Turnover Days', format14)
        sheet.merge_range(17, 0, 17, 2, 'Avg Payable', format14)
        sheet.merge_range(18, 0, 18, 2, 'Operating  Cycle', format14)
        sheet.merge_range(4,3,4,4,'SR', format11)
        
       
        if current_result:
            sheet.write('D4',''+ str(convert_date.day) + '/' + str(convert_date.year), format11)
            for result in current_result:
                sheet.merge_range(5,3,5,4,result['gross_profit_margin'],format12)
                sheet.merge_range(6,3,6,4,result['net_profit_percentage'],format12)
                sheet.merge_range(7,3,7,4,result['current_ratio'],format12)
                sheet.merge_range(8,3,8,4,result['net_working_capital'],format12)
                sheet.merge_range(9,3,9,4,result['debit_equity_ratio'],format12)
                sheet.merge_range(10,3,10,4,result['equity_asset_ratio'],format12)
                sheet.merge_range(11,3,11,4,result['shareholder_equity'],format13)
                sheet.merge_range(12,3,12,4,result['receivable_turnover'],format12)
                sheet.merge_range(13,3,13,4,result['receivable_turnover_in_day'],format12)
                sheet.merge_range(14,3,14,4,result['average_receivable'],format13)
                sheet.merge_range(15,3,15,4,result['payable_turnover'],format12)
                sheet.merge_range(16,3,16,4,result['payable_turnover_days'],format12)
                sheet.merge_range(17,3,17,4,result['average_payable'],format13)
                sheet.merge_range(18,3,18,4,result['operating_cycle'],format12)


RatioReportXlsx('report.ratio_analysis.ratio_analysis_report_xls.xlsx', 'ratio.analysis') 


