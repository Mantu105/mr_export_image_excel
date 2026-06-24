# -*- coding: utf-8 -*-
import base64
import datetime
import io
import json
import logging

from odoo import _
from odoo.exceptions import UserError
from odoo.http import request
from odoo.tools.image import image_process

from odoo.addons.web.controllers.export import ExcelExport

_logger = logging.getLogger(__name__)

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None

# Visual sizing of the embedded thumbnails.
IMAGE_MAX_PX = 90          # max width/height of the embedded thumbnail (pixels)
IMAGE_ROW_HEIGHT = 74      # row height in points (~90px + small padding)
IMAGE_COL_WIDTH = 16       # column width in Excel character units
CONFIG_PARAM = 'mr_export_image_excel.export_image_in_excel'


class ExcelExportImage(ExcelExport):
    """Extend the standard XLSX exporter so that binary / image columns are
    embedded as real images in the generated file instead of base64 text.

    The behaviour is controlled by the *Export Image in Excel* option in the
    General Settings (Permissions). When the option is disabled the standard
    Odoo behaviour is preserved.
    """

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _evo_image_export_enabled(self):
        """Return True when the feature is enabled in the settings."""
        try:
            value = request.env['ir.config_parameter'].sudo().get_param(CONFIG_PARAM)
        except Exception:  # pragma: no cover - extremely defensive
            return False
        return value in ('True', 'true', '1', 1, True)

    def _evo_resolve_field(self, model_obj, path):
        """Resolve a (possibly dotted) export path to its final ``Field``."""
        field = None
        current = model_obj
        for part in (path or '').split('/'):
            if part in ('id', '.id'):
                return current._fields.get('id')
            field = current._fields.get(part)
            if field is None:
                return None
            if field.relational and field.comodel_name:
                current = request.env[field.comodel_name]
        return field

    def _evo_compute_image_columns(self, model, fields):
        """Return a set of column indexes that point to a binary/image field."""
        columns = set()
        if not model or model not in request.env:
            return columns
        model_obj = request.env[model]
        for index, descriptor in enumerate(fields or []):
            name = descriptor.get('name') if isinstance(descriptor, dict) else None
            if not name:
                continue
            try:
                field = self._evo_resolve_field(model_obj, name)
            except Exception:
                field = None
            if field is not None and field.type == 'binary':
                columns.add(index)
        return columns

    def _evo_insert_image(self, worksheet, row, col, cell_value, cell_format):
        """Insert ``cell_value`` (base64 image) into the worksheet cell.

        Returns True if an image was inserted, False otherwise (caller then
        falls back to the default text rendering).
        """
        if not cell_value:
            return False
        try:
            # Odoo export returns base64-encoded strings; decode to raw bytes first.
            if isinstance(cell_value, str):
                raw = base64.b64decode(cell_value)
            elif isinstance(cell_value, bytes):
                raw = base64.b64decode(cell_value)
            else:
                return False

            # image_process expects and returns raw binary bytes (not base64).
            thumbnail = image_process(raw, size=(IMAGE_MAX_PX, IMAGE_MAX_PX))
            if not thumbnail:
                return False

            worksheet.write_blank(row, col, None, cell_format)
            worksheet.insert_image(row, col, 'image.png', {
                'image_data': io.BytesIO(thumbnail),
                'x_offset': 3,
                'y_offset': 3,
                'object_position': 2,  # move but don't size with cells
            })
            return True
        except Exception:
            _logger.debug("mr_export_image_excel: could not embed image at "
                          "row %s col %s, falling back to text.", row, col,
                          exc_info=True)
            return False

    # ------------------------------------------------------------------
    # Overrides
    # ------------------------------------------------------------------
    def base(self, data, *args, **kwargs):
        """Capture which exported columns are binary/image fields so that
        :meth:`from_data` can embed them as images.
        """
        self._evo_image_cols = set()
        try:
            params = json.loads(data)
            self._evo_image_cols = self._evo_compute_image_columns(
                params.get('model'), params.get('fields') or [])
        except Exception:
            self._evo_image_cols = set()
        return super().base(data, *args, **kwargs)

    def from_data(self, fields, columns_headers, rows):
        image_cols = getattr(self, '_evo_image_cols', set()) or set()
        export_images = bool(image_cols) and self._evo_image_export_enabled()

        # If the feature is off or there is no image column, keep the default
        # Odoo behaviour untouched.
        if not export_images or xlsxwriter is None:
            return super().from_data(fields, columns_headers, rows)

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet()

        if len(rows) > worksheet.xls_rowmax:
            raise UserError(_(
                "There are too many rows (%(count)s rows, limit: %(limit)s) to "
                "export as an Excel file. Consider splitting the export.",
                count=len(rows), limit=worksheet.xls_rowmax,
            ))

        header_style = workbook.add_format({
            'bold': True, 'bg_color': '#D9D9D9', 'border': 1,
            'align': 'center', 'valign': 'vcenter',
        })
        base_style = workbook.add_format({'valign': 'vcenter'})
        date_style = workbook.add_format(
            {'num_format': 'yyyy-mm-dd', 'valign': 'vcenter'})
        datetime_style = workbook.add_format(
            {'num_format': 'yyyy-mm-dd hh:mm:ss', 'valign': 'vcenter'})
        image_style = workbook.add_format(
            {'align': 'center', 'valign': 'vcenter'})

        # Header row
        for col, header in enumerate(columns_headers):
            worksheet.write(0, col, header, header_style)
            if col in image_cols:
                worksheet.set_column(col, col, IMAGE_COL_WIDTH)
            else:
                worksheet.set_column(col, col, min(max(len(header) * 1.1, 10), 50))

        # Data rows
        for row_index, row in enumerate(rows):
            target_row = row_index + 1
            row_has_image = False
            for col, cell_value in enumerate(row):
                if col in image_cols and cell_value:
                    if self._evo_insert_image(worksheet, target_row, col,
                                              cell_value, image_style):
                        row_has_image = True
                        continue

                # --- default value rendering ------------------------------
                if isinstance(cell_value, (list, tuple)):
                    cell_value = str(cell_value)
                if isinstance(cell_value, bytes):
                    try:
                        cell_value = cell_value.decode('utf-8')
                    except UnicodeDecodeError:
                        raise UserError(_(
                            "Binary fields can not be exported to Excel unless "
                            "their content is base64-encoded. That does not seem "
                            "to be the case for %s.", columns_headers[col]))
                if isinstance(cell_value, str):
                    if len(cell_value) > 32767:
                        cell_value = _(
                            "The content of this cell is too long for an XLSX "
                            "file (more than %s characters). Please use the CSV "
                            "format for this export.", 32767)
                    else:
                        cell_value = cell_value.replace("\r", " ")
                    worksheet.write_string(target_row, col, cell_value or '', base_style)
                elif isinstance(cell_value, datetime.datetime):
                    worksheet.write_datetime(target_row, col, cell_value, datetime_style)
                elif isinstance(cell_value, datetime.date):
                    worksheet.write_datetime(target_row, col, cell_value, date_style)
                elif cell_value is None or cell_value is False:
                    worksheet.write_blank(target_row, col, None, base_style)
                else:
                    worksheet.write(target_row, col, cell_value, base_style)

            if row_has_image:
                worksheet.set_row(target_row, IMAGE_ROW_HEIGHT)

        workbook.close()
        return output.getvalue()
