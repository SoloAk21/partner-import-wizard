from odoo import models, fields

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