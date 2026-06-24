# -*- coding: utf-8 -*-
from odoo import fields, models


class ResConfigSettings(models.TransientModel):
    _inherit = 'res.config.settings'

    export_image_in_excel = fields.Boolean(
        string="Export Image in Excel",
        config_parameter='mr_export_image_excel.export_image_in_excel',
        help="When enabled, binary / image columns selected during an Excel "
             "(XLSX) export are embedded as images in the generated file "
             "instead of base64 text.",
    )
