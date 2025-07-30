from odoo import models, fields, api, _
import base64
import csv
import io
from odoo.exceptions import UserError, ValidationError
import logging

# For Excel support
try:
    import openpyxl
    EXCEL_SUPPORT = True
except ImportError:
    EXCEL_SUPPORT = False
    logging.warning("openpyxl library not found. Excel import will be disabled.")

class PartnerImportWizard(models.TransientModel):
    _name = 'partner.import.wizard'
    _description = 'Partner Import Wizard'

    file = fields.Binary(string='Upload File', required=True)
    file_name = fields.Char(string='File Name')
    import_mode = fields.Selection([
        ('create', 'Create New Records'),
        ('update', 'Update Existing Records'),
        ('both', 'Create and Update'),
    ], string='Import Mode', default='create', required=True)
    file_type = fields.Selection([
        ('csv', 'CSV'),
        ('xlsx', 'Excel (XLSX)'),
    ], string='File Type', compute='_compute_file_type', store=True)

    @api.depends('file_name')
    def _compute_file_type(self):
        for record in self:
            if record.file_name:
                if record.file_name.lower().endswith('.xlsx'):
                    record.file_type = 'xlsx'
                elif record.file_name.lower().endswith('.csv'):
                    record.file_type = 'csv'
                else:
                    record.file_type = False
            else:
                record.file_type = 'csv'

    @api.constrains('file_name')
    def _check_file_type(self):
        for record in self:
            if record.file_name and not record.file_type:
                raise ValidationError(_(
                    "Unsupported file format. Please upload a CSV or Excel (XLSX) file."
                ))

    def process_file(self):
        self.ensure_one()
        
        if not self.file_name:
            raise UserError(_('Please upload a file.'))
            
        if not self.file_type:
            raise UserError(_(
                "Unsupported file format '%s'. Please upload a CSV or Excel (XLSX) file."
            ) % self.file_name.split('.')[-1])
            
        if self.file_type == 'xlsx' and not EXCEL_SUPPORT:
            raise UserError(_(
                "Excel import requires the openpyxl library. "
                "Please install it with: pip install openpyxl"
            ))
        
        try:
            file_content = base64.b64decode(self.file)
            
            if self.file_type == 'csv':
                result = self._process_csv(file_content)
            else:
                result = self._process_excel(file_content)
                
            return result
                
        except Exception as e:
            error_msg = _("Failed to process file: %s") % str(e)
            logging.error(error_msg, exc_info=True)
            raise UserError(error_msg)

    def _process_csv(self, file_content):
        """Process CSV file content"""
        try:
            # Try UTF-8 first, fallback to other encodings if needed
            try:
                file_string = file_content.decode('utf-8-sig')  # Handle BOM if present
            except UnicodeDecodeError:
                # Try common alternative encodings
                for encoding in ['latin-1', 'iso-8859-1', 'cp1252']:
                    try:
                        file_string = file_content.decode(encoding)
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    raise UserError(_(
                        "Could not decode the file. Please ensure it's a valid CSV file with "
                        "UTF-8, Latin-1, or Windows-1252 encoding."
                    ))
            
            file_io = io.StringIO(file_string)
            reader = csv.DictReader(file_io)
            return self._process_rows(reader)
            
        except csv.Error as e:
            raise UserError(_(
                "Invalid CSV file format. Please check the file structure. Error: %s"
            ) % str(e))
        except Exception as e:
            raise UserError(_("Failed to process CSV file: %s") % str(e))

    def _process_excel(self, file_content):
        """Process Excel file content"""
        try:
            excel_file = io.BytesIO(file_content)
            workbook = openpyxl.load_workbook(excel_file, read_only=True)
            sheet = workbook.active
            
            headers = []
            for cell in sheet[1]:
                headers.append(str(cell.value).lower() if cell.value else '')
            
            rows = []
            for row in sheet.iter_rows(min_row=2):
                row_data = {}
                for idx, cell in enumerate(row):
                    if idx < len(headers) and headers[idx]:
                        row_data[headers[idx]] = cell.value
                if any(val for val in row_data.values() if val not in (None, "")):
                    rows.append(row_data)
            
            return self._process_rows(rows)
            
        except Exception as e:
            raise UserError(_(
                "Invalid Excel file. Please ensure it's a valid XLSX file. Error: %s"
            ) % str(e))

    def _process_rows(self, rows):
        """Common processing for both CSV and Excel rows"""
        created = 0
        updated = 0
        errors = []
        
        for idx, row in enumerate(rows, start=2):  # Row numbers start at 2 (1 is header)
            try:
                if not any(val for val in row.values() if val not in (None, "")):
                    continue
                    
                if not row.get('name') or not row.get('email'):
                    errors.append(_(
                        "Row %d: Missing required field - Name: %s, Email: %s"
                    ) % (idx, row.get('name', ''), row.get('email', '')))
                    continue
                
                partner_vals = {
                    'name': row['name'],
                    'email': row['email'],
                    'phone': row.get('phone'),
                    'street': row.get('street'),
                    'city': row.get('city'),
                    'zip': row.get('zip'),
                    'country_id': self._get_country(row.get('country')),
                }
                
                partner = self.env['res.partner'].search([('email', '=', row['email'])], limit=1)
                
                if partner and self.import_mode in ('update', 'both'):
                    partner.write(partner_vals)
                    updated += 1
                elif not partner and self.import_mode in ('create', 'both'):
                    self.env['res.partner'].create(partner_vals)
                    created += 1
                else:
                    errors.append(_(
                        "Row %d: Skipped - %s (Import mode doesn't allow this operation)"
                    ) % (idx, row['email']))
            
            except Exception as e:
                errors.append(_("Row %d: Error processing - %s") % (idx, str(e)))
        
        # Prepare notification based on results
        base_message = _(
            "Import completed: %(created)d created, %(updated)d updated",
            created=created,
            updated=updated
        )
        
        if errors:
            error_count = len(errors)
            error_samples = "\n".join(errors[:3])  # Show first 3 errors as samples
            if error_count > 3:
                error_samples += _("\n...and %d more errors") % (error_count - 3)
            
            notification = {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': _('Import Completed with Errors'),
                    'message': f"{base_message}\n\n"
                              f"{_('Errors encountered:')} {error_count}\n"
                              f"{error_samples}",
                    'sticky': True,
                    'type': 'warning',
                    'next': {
                        'type': 'ir.actions.act_window_close'
                    }
                }
            }
        else:
            notification = {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': _('Import Successful'),
                    'message': base_message,
                    'sticky': False,
                    'type': 'success',
                    'next': {
                        'type': 'ir.actions.act_window_close'
                    }
                }
            }
        
        # Also log full results for admin review
        log_message = f"Partner Import Results:\n{base_message}\n"
        if errors:
            log_message += f"Errors:\n" + "\n".join(errors)
        logging.info(log_message)
        
        return notification

    def _get_country(self, country_name):
        if not country_name:
            return False
        country = self.env['res.country'].search([('name', '=', country_name)], limit=1)
        return country.id if country else False