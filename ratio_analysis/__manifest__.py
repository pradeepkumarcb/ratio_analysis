# -*- coding: utf-8 -*-
##############################################################################
#
#    ODOO, Open Source Management Solution
#    Copyright (C) 2018-TODAY GSIT
#    For more details, check COPYRIGHT and LICENSE files
#
##############################################################################
{
    'name': 'GSIT Ratio Analysis',
    'category': 'Reporting',
    'summary': 'module Ratio Analysis Report',
    'version': '10.0.1.0.1',
    'application': False,
    'author': 'GSIT',
    'depends': [
                'account_reports_extended', 
                ],
    'license': 'LGPL-3',
    'data': ['security/ir.model.access.csv',
            'wizard/gs_ratio_analysis_view.xml',
            'view/ratio_analysis_excel.xml',
             ],
    
    'installable': True,
    'auto_install': False
    
}
