# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import api, models, fields, _
import xlsxwriter
from cStringIO import StringIO
import base64
import datetime

class WizardValuationHistory(models.TransientModel):
    _inherit = 'wizard.valuation.history'
    _description = 'Wizard that opens the stock valuation history table'

    date = fields.Datetime('Date', default=datetime.datetime.today(), required=True)
    filename = fields.Char('Filename')
    document = fields.Binary(string = 'Download Excel')
    need_export = fields.Boolean(string="Export Excel?", default=False)

    @api.multi
    def open_table(self):
        self.ensure_one()
        cr= self._cr
        if self.need_export==True:
            Header_Text = 'Inventory_at_date'
            file_data = StringIO()
            workbook = xlsxwriter.Workbook(file_data)
            worksheet = workbook.add_worksheet("Inventory")
            bold = workbook.add_format({'bold': 1,'align':'center','border':1})
            header_format = workbook.add_format({'bold': 1,'align':'center','valign':'vcenter','fg_color':'yellow','border':1})
            worksheet.set_column('A3:A3',22)
            worksheet.set_column('B3:B3',30)
            worksheet.set_column('C3:C3',20)
            worksheet.set_column('D3:C3',20)
            worksheet.set_column('E3:E3',30)
            worksheet.set_column('F3:F3',10)
            worksheet.set_column('G3:G3',15)
            worksheet.set_column('H3:H3',15)
            #Header configuration
            preview = 'Inventory At Date Report:'+str(self.date or datetime.datetime.today())

            worksheet.merge_range('C1:F2',preview, header_format)

            for i in range(1):
                worksheet.write('A3', 'Product',bold)
                worksheet.write('B3', 'Location',bold)
                worksheet.write('C3', 'Company',bold)
                worksheet.write('D3', 'Operation Date',bold)
                worksheet.write('E3', 'Move',bold)
                worksheet.write('F3', 'Source',bold)
                worksheet.write('G3', 'Product Quantity',bold)
                worksheet.write('H3', 'Inventory Value',bold)
                
            cr.execute("""
                        SELECT 
                                move_id,
                                location_id,
                                company_id,
                                product_id,
                                product_categ_id,
                                product_name,
                                product_template_id,
                                SUM(quantity) as quantity,
                                date,
                                COALESCE(SUM(price_unit_on_quant * quantity) / NULLIF(SUM(quantity), 0), 0) as price_unit_on_quant,
                                source,
                                string_agg(DISTINCT serial_number, ', ' ORDER BY serial_number) AS serial_number
                                FROM
                                ((SELECT
                                    stock_move.id AS id,
                                    stock_move.name AS move_id,
                                    dest_location.complete_name AS location_id,
                                    res_company.name AS company_id,
                                    product_product.default_code AS product_id,
                                    product_template.name AS product_name,
                                    product_template.id AS product_template_id,
                                    product_template.categ_id AS product_categ_id,
                                    quant.qty AS quantity,
                                    stock_move.date AS date,
                                    quant.cost as price_unit_on_quant,
                                    stock_move.origin AS source,
                                    stock_production_lot.name AS serial_number
                                FROM
                                    stock_quant as quant
                                JOIN
                                    stock_quant_move_rel ON stock_quant_move_rel.quant_id = quant.id
                                JOIN
                                    stock_move ON stock_move.id = stock_quant_move_rel.move_id
                                LEFT JOIN
                                    stock_production_lot ON stock_production_lot.id = quant.lot_id
                                JOIN
                                    stock_location dest_location ON stock_move.location_dest_id = dest_location.id
                                JOIN
                                    stock_location source_location ON stock_move.location_id = source_location.id
                                JOIN
                                    product_product ON product_product.id = stock_move.product_id
                                JOIN
                                    product_template ON product_template.id = product_product.product_tmpl_id
                                JOIN
                                    res_company ON dest_location.company_id = res_company.id
                                WHERE quant.qty>0 AND stock_move.state = 'done' AND dest_location.usage in ('internal', 'transit') AND stock_move.date <=%s
                                AND (
                                    not (source_location.company_id is null and dest_location.company_id is null) or
                                    source_location.company_id != dest_location.company_id or
                                    source_location.usage not in ('internal', 'transit'))
                                ) UNION ALL
                                (SELECT
                                    (-1) * stock_move.id AS id,
                                    stock_move.name AS move_id,
                                    source_location.complete_name AS location_id,
                                    res_company.name AS company_id,
                                    product_product.default_code AS product_id,
                                    product_template.name AS product_name,
                                    product_template.id AS product_template_id,
                                    product_template.categ_id AS product_categ_id,
                                    - quant.qty AS quantity,
                                    stock_move.date AS date,
                                    quant.cost as price_unit_on_quant,
                                    stock_move.origin AS source,
                                    stock_production_lot.name AS serial_number
                                FROM
                                    stock_quant as quant
                                JOIN
                                    stock_quant_move_rel ON stock_quant_move_rel.quant_id = quant.id
                                JOIN
                                    stock_move ON stock_move.id = stock_quant_move_rel.move_id
                                LEFT JOIN
                                    stock_production_lot ON stock_production_lot.id = quant.lot_id
                                JOIN
                                    stock_location source_location ON stock_move.location_id = source_location.id
                                JOIN
                                    stock_location dest_location ON stock_move.location_dest_id = dest_location.id
                                JOIN
                                    product_product ON product_product.id = stock_move.product_id
                                JOIN
                                    product_template ON product_template.id = product_product.product_tmpl_id
                                JOIN
                                    res_company ON source_location.company_id = res_company.id
                                WHERE quant.qty>0 AND stock_move.state = 'done' AND source_location.usage in ('internal', 'transit') AND stock_move.date <=%s
                                AND (
                                    not (dest_location.company_id is null and source_location.company_id is null) or
                                    dest_location.company_id != source_location.company_id or
                                    dest_location.usage not in ('internal', 'transit'))
                                ))
                                AS foo
                                GROUP BY move_id, location_id, company_id, product_id, product_name,product_categ_id, date, source, product_template_id
                        """,(str(self.date or datetime.datetime.today()),str(self.date or datetime.datetime.today()),))
            row=4
            col=0
            
            for each_line in cr.dictfetchall():
                worksheet.write(row,col , each_line.get('product_id') or each_line.get('product_name')or '')
                worksheet.write(row,col+1 ,each_line.get('location_id') or '')
                worksheet.write(row,col+2 ,each_line.get('company_id') or '')
                worksheet.write(row,col+3 ,each_line.get('date'))
                worksheet.write(row,col+4 ,each_line.get('move_id') or '')
                worksheet.write(row,col+5 ,each_line.get('source') or '')
                worksheet.write(row,col+6 ,each_line.get('quantity') )
                worksheet.write(row,col+7 ,each_line.get('price_unit_on_quant')*each_line.get('quantity'))
                row+=1

            workbook.close()
            file_data.seek(0)
            self.write({'document':base64.encodestring(file_data.read()), 'filename':Header_Text+'.xlsx'})
            return {
            'name': _('Inventory Value At Date'),
            'res_model':'wizard.valuation.history',
            'type':'ir.actions.act_window',
            'view_type':'form',
            'view_mode':'form',
            'target':'new',
            'nodestroy': True,
            # 'context': context,
            'res_id': self.id
            }

        ctx = dict(
            self._context,
            history_date=self.date,
            search_default_group_by_product=False,
            search_default_group_by_location=False)
        return {
            'domain': "[('date', '<=', '" + self.date + "')]",
            'name': _('Inventory Value At Date'),
            'view_type': 'form',
            'view_mode': 'tree',
            'res_model': 'stock.history',
            'type': 'ir.actions.act_window',
            'context': ctx,
        }

