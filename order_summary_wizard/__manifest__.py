{
    'name': 'Order Summary Wizard',
    'version': '1.0',
    'depends': ['sale', 'stock'],
    'external_dependencies': {'python': ['openpyxl']},
    'data': [
        'security/ir.model.access.csv',
        'views/order_summary_wizard_views.xml',
        'reports/order_summary_report.xml',
        'reports/order_summary_report_template.xml',
    ],
    'assets': {
        'web.assets_backend': [
            'order_summary_wizard/static/src/css/styles.css',
            'order_summary_wizard/static/src/js/custom_script.js',
        ],
    },
    'installable': True,
    'application': False,
}