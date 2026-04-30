{
    'name': 'Attachment files preview',
    'version': '17.0',
    'depends': ['web', 'base', 'mail', ],
    'author': 'srikanth',
    'description': """ CSV ,XLSX ,PPTX AND  DOCX Attachments files preview """,
    'assets': {
        'web.assets_backend': [
            # preview docs
            'ts_office_files_preview/static/lib/jszip.min.js',
            # 'ts_office_files_preview/static/lib/dock-preview.js',

            'ts_office_files_preview/static/src/js/chatter_preview.js',
            'ts_office_files_preview/static/src/css/document_preview.css',

            # preview xls and csv
            "ts_office_files_preview/static/lib/xls/spreadheet.js",
            "ts_office_files_preview/static/src/css/spreadhseet.css",
            "ts_office_files_preview/static/lib/xls/xlsx.full.min.js",

            # preview pptx
            "ts_office_files_preview/static/lib/pdf.mjs",
            "ts_office_files_preview/static/lib/pdf.worker.mjs",
        ],
    },
}
