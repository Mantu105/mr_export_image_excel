# -*- coding: utf-8 -*-
{
    'name': 'Export Image in Excel',
    'version': '18.0.1.0.0',
    'category': 'Extra Tools',
    'summary': 'Directly export images in Excel using the standard Odoo export feature.',
    'description': """
Export Image in Excel
=====================

Using this module, you can directly export images in Excel using the standard
Odoo export feature.

When you export records (for example Products) to XLSX and include a binary /
image column (such as the product image) in the export list, the picture is
embedded as a real image inside the generated Excel file instead of a long
base64 text string.

How to use
----------
* Go to **Settings -> General Settings -> Permissions -> Export Image in Excel**
  and enable the option.
* Select some records (e.g. Products) and use **Action -> Export**.
* Add the image column (e.g. *Image*) to the list of fields to export and
  export in **xlsx** format.
* Open the Excel file: the images are shown directly in the file.

Key Features
------------
* Export image directly in Excel.
* Works with any model that has an image/binary field.
* Easy to use, no extra Python library required.
""",
    'author': 'MR',
    'website': '',
    'license': 'LGPL-3',
    'depends': [
        'web',
        'base_setup',
    ],
    'data': [
        'views/res_config_settings_views.xml',
    ],
    'images': ['static/description/banner.png'],
    'installable': True,
    'application': False,
    'auto_install': False,
}
