# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

{
    'name': 'Export Inventory Report to Excel',
    'category': 'warehouse',
    'description': """
A module that provide support to handle the large inventory records.
=====================================

This module is usefull to export the inventory records into excel and avoid the slowness with large inventory data. """,
    'depends': ['base', 'stock','stock_account'],
    'data': [
        'views/inventory_view.xml',
    ],
    'author':'arpit',
    'installable': True,
    'application': True,
}
