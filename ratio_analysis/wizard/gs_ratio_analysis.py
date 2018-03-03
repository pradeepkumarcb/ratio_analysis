# -*- coding: utf-8 -*-
##############################################################################
#
#    ODOO, Open Source Management Solution
#    Copyright (C) 2016-TODAY Steigend IT Solutions
#    For more details, check COPYRIGHT and LICENSE files
#
##############################################################################
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from odoo import api, fields, models


class RatioAnalysis(models.Model):
    _name = 'ratio.analysis'
    
    from_date = fields.Date('From Date')
    to_date = fields.Date('To Date')
    enable_cmp = fields.Boolean('Enable Comparison With')
    cmp_date_from = fields.Date('From Date')
    cmp_date_to = fields.Date('To Date')
    
    @api.multi
    def export_xls(self):
        context = self._context
        
        datas = {'ids': context.get('active_ids', [])}
        datas['model'] = 'ratio.analysis'
        datas['form'] = self.read()[0]
        for field in datas['form'].keys():
            if isinstance(datas['form'][field], tuple):
                datas['form'][field] = datas['form'][field][0]
        if context.get('xls_export'):
            return {'type': 'ir.actions.report.xml',
                    'report_name': 'ratio_analysis.ratio_analysis_report_xls.xlsx',
                    'datas': datas,
                    'name': 'Ratio Analysis Report'
                    }
    
    
    
    
    