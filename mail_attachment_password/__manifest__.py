# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

{
    'name': 'Set Password on Email Document',
    'category': 'Enhancement',
    'summary':'Easily protect the template attachment (PDF/Excel) by setting password.',
    'description': """
A module that provide support to set password on email template document.
=====================================

This module gives the opportunity to make the reports(pdf/excel attachment) protected while sending email """,
    'depends': ['base', 'mail'],
    'data': [
        'views/template_attachment_password.xml',
    ],
    'author':'arpit',
    'installable': True,
    'application': True,
    'images':['static/description/Banner.jpg']
}
