# -*- coding: utf-8 -*-
from odoo import models, api


class CollectionReportPDF(models.AbstractModel):
    """
    QWeb PDF report parser.
    The report is printed FROM the wizard record so `docs` is a wizard
    recordset (single record).
    """
    _name = 'report.collection_report.collection_report_template'
    _description = 'Collection Report PDF Parser'

    @api.model
    def _get_report_values(self, docids, data=None):
        wizards = self.env['collection.report.wizard'].browse(docids)
        report_lines = []
        for wiz in wizards:
            lines = wiz._get_report_data()
            # ── group by city ──────────────────────────────────
            cities = {}
            for line in lines:
                city = line['city'] or 'Unknown'
                cities.setdefault(city, []).append(line)

            city_summaries = []
            grand_total = 0.0
            grand_residual = 0.0
            for city, city_lines in sorted(cities.items()):
                total = sum(l['amount_total'] for l in city_lines)
                residual = sum(l['amount_residual'] for l in city_lines)
                grand_total += total
                grand_residual += residual
                city_summaries.append({
                    'city': city,
                    'lines': city_lines,
                    'total': total,
                    'residual': residual,
                    'count': len(city_lines),
                })

            report_lines.append({
                'wizard': wiz,
                'city_summaries': city_summaries,
                'grand_total': grand_total,
                'grand_residual': grand_residual,
                'total_invoices': len(lines),
            })

        return {
            'doc_ids': docids,
            'doc_model': 'collection.report.wizard',
            'docs': wizards,
            'report_data': report_lines,
        }
