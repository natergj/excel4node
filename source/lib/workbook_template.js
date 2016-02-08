let utils = require('./utils.js');

module.exports = {
    'workbook': {
        '@xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        '@xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'fileSharing': {},
        'bookViews': [
            {
                'workbookView': {
                    '@tabRatio': '600',
                    '@windowHeight': '14980',
                    '@windowWidth': '25600',
                    '@xWindow': '0',
                    '@yWindow': '1080'
                }
            }
        ],
        'sheets': [],
        'definedNames': []
    },
    'strings': {
        'sst': [
            {
                '@count': 0,
                '@uniqueCount': 0,
                '@xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            }
        ]
    },
    'workbook_xml_rels': {
        'Relationships': [
            {
                '@xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'
            },
            {
                'Relationship': {
                    '@Id': utils.generateRId(),
                    '@Target': 'sharedStrings.xml',
                    '@Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
                }
            },
            {
                Relationship: {
                    '@Id': utils.generateRId(),
                    '@Target': 'styles.xml',
                    '@Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
                }
            }
        ]
    },
    global_rels: {
        Relationships: [
            {
                '@xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'
            },
            {
                Relationship: {
                    '@Id': utils.generateRId(),
                    '@Target': 'xl/workbook.xml',
                    '@Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
                }
            }
        ]
    },
    Content_Types: {
        Types: [
            {
                '@xmlns': 'http://schemas.openxmlformats.org/package/2006/content-types'
            },
            {
                Default: {
                    '@ContentType': 'application/xml',
                    '@Extension': 'xml'
                }
            },
            {
                Default: {
                    '@ContentType': 'application/vnd.openxmlformats-package.relationships+xml',
                    '@Extension': 'rels'
                }
            },
            {
                Override: {
                    '@ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                    '@PartName': '/xl/workbook.xml'
                }
            },
            {
                Override: {
                    '@ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
                    '@PartName': '/xl/styles.xml'
                }
            },
            {
                Override: {
                    '@ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
                    '@PartName': '/xl/sharedStrings.xml'
                }
            }
        ]
    },
    sharedStrings: [],
    debug: false
};