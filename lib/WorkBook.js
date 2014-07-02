var ws = require('./WorkSheet.js'),
style = require('./Style.js'),
xml = require('xmlbuilder'),
jszip = require('jszip'),
fs = require('fs');

exports.WorkBook = function(){
	this.workbook = {
		WorkSheets:[],
		workbook:{
			'@xmlns:r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
			'@xmlns':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
			sheets:[]
		},
		styles:{
			styleSheet:{
				'@xmlns':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
				'@mc:Ignorable':'x14ac',
				'@xmlns:mc':'http://schemas.openxmlformats.org/markup-compatibility/2006',
				'@xmlns:x14ac':'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
				numFmts:[
					{
						'@count':0
					}
				],
				fonts:[
					{
						'@count':1
					},
					{
						font:[
							{
								sz:{
									'@val':12
								}
							},
							{
								color:{
									'@theme':1
								}
							},
							{
								name:{
									'@val':'Calibri'
								}
							},
							{
								family:{
									'@val':'2'
								}
							},
							{
								scheme:{
									'@val':'minor'
								}
							}
						]
					}
				],
				fills:[
					{
						'@count':2
					},
					{    
						fill:{
							patternFill:{
								'@patternType':'none'
							}
						}
					},
					{    
						fill:{
							patternFill:{
								'@patternType':'gray125'
							}
						}
					}
				],
				borders:[
					{
						'@count':1
					},
					{
						border:[
							'left',
							'right',
							'top',
							'bottom',
							'diagonal'
      					]
					}
				],
				cellStyleXfs:[
					{
						'@count':1
					},
					{
						xf:{
							'@borderId':0,
							'@fillId':0,
							'@fontId':0,
							'@numFmtId':0
						}
					}
				],
				cellXfs:[
					{
						'@count':1
					},
					{
						xf:{
							'@applyNumberFormat':0,
							'@borderId':0,
							'@fillId':0,
							'@fontId':0,
							'@numFmtId':164,
							'@xfId':0
						} 
					}
				]
			}
		},
		strings : {
			sst:[
				{
					'@count':0,
					'@uniqueCount':0,
					'@xmlns':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
				}
			]
		},
		workbook_xml_rels:{
			Relationships:[
				{
					'@xmlns':'http://schemas.openxmlformats.org/package/2006/relationships'
				},
				{
					Relationship:{
						'@Id':generateRId(),
						'@Target':'sharedStrings.xml',
						'@Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
					}
				},
				{
					Relationship:{
						'@Id':generateRId(),
						'@Target':'styles.xml',
						'@Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
					}
				}
			]
		},
		global_rels:{
			Relationships:[
				{
					'@xmlns':'http://schemas.openxmlformats.org/package/2006/relationships'
				},
				{
					Relationship:{
						'@Id':generateRId(),
						'@Target':'xl/workbook.xml',
						'@Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
					}
				}
			]
		},
		Content_Types:{
			Types:[
				{
					'@xmlns':'http://schemas.openxmlformats.org/package/2006/content-types'
				},
				{
					Default:{
						'@ContentType':'application/xml',
						'@Extension':'xml'
					}
				},
				{
					Default:{
						'@ContentType':'application/vnd.openxmlformats-package.relationships+xml',
						'@Extension':'rels'
					} 
				},
				{
					Override:{
						'@ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
						'@PartName':'/xl/workbook.xml'
					}
				},
				{
					Override:{
						'@ContentType':'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
						'@PartName':'/xl/styles.xml'
					}
				},
				{
					Override:{
						'@ContentType':'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
						'@PartName':'/xl/sharedStrings.xml'
					}
				}
			]
		},
		sharedStrings:[],
		sharedStyles:{
			numFmts:[],
			fonts:[],
			fills:[],
			borders:[],
			xf:[]
		}
	}

	this.write = function(fileName, response){

		var xlsx = new jszip();
		var that = this;
		var sheetCount = 1;
		this.workbook.WorkSheets.forEach(function(sheet){
			var wsObj = {'worksheet':JSON.parse(JSON.stringify(sheet.sheet))};
			var xmlStr = xml.create(wsObj);
			var thisRId = generateRId();
			that.workbook.workbook_xml_rels.Relationships.push({
				Relationship:{
					'@Id':thisRId,
					'@Target':'worksheets/sheet'+sheetCount+'.xml',
					'@Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
				}
			})
			that.workbook.workbook.sheets.push({
				sheet:{
					'@name':sheet.name,
					'@sheetId':sheetCount,
					'@r:id':thisRId
				}
			});			
			that.workbook.Content_Types.Types.push({
				Override:{
					'@ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
					'@PartName':'/xl/worksheets/sheet'+sheetCount+'.xml'
				}
			});
			//console.log(xmlStr.end({ pretty: true, indent: '  ', newline: '\n' }));

			xlsx.folder("xl").folder("worksheets").file('sheet'+sheetCount+'.xml',xmlStr.end({ pretty: true, indent: '  ', newline: '\n' }));
			sheetCount+=1;
		});

		this.workbook.sharedStrings.forEach(function(s){
			that.workbook.strings.sst.push({'si':{'t':s}});
		});
		this.workbook.strings.sst[0]['@uniqueCount']=this.workbook.sharedStrings.length;




		var wbXML = xml.create({workbook:JSON.parse(JSON.stringify(this.workbook.workbook))});
		//console.log(wbXML.end({ pretty: true, indent: '  ', newline: '\n' }));

		var styleXML = xml.create(JSON.parse(JSON.stringify(this.workbook.styles)));
		console.log(styleXML.end({ pretty: true, indent: '  ', newline: '\n' }));

		var relsXML = xml.create(this.workbook.workbook_xml_rels);
		//console.log(relsXML.end({ pretty: true, indent: '  ', newline: '\n' }));

		var Content_TypesXML = xml.create(this.workbook.Content_Types);
		//console.log(Content_TypesXML.end({ pretty: true, indent: '  ', newline: '\n' }));

		var globalRelsXML = xml.create(this.workbook.global_rels);
		//console.log(globalRelsXML.end({ pretty: true, indent: '  ', newline: '\n' }));

		var stringsXML = xml.create(this.workbook.strings);
		//console.log(stringsXML.end({ pretty: true, indent: '  ', newline: '\n' }));


		xlsx.file("[Content_Types].xml",Content_TypesXML.end({ pretty: true, indent: '  ', newline: '\n' }));
		xlsx.folder("_rels").file(".rels",globalRelsXML.end({ pretty: true, indent: '  ', newline: '\n' }));
		xlsx.folder("xl").file("sharedStrings.xml",stringsXML.end({ pretty: true, indent: '  ', newline: '\n' }));
		xlsx.folder("xl").file("styles.xml",styleXML.end({ pretty: true, indent: '  ', newline: '\n' }));
		xlsx.folder("xl").file("workbook.xml",wbXML.end({ pretty: true, indent: '  ', newline: '\n' }));
		xlsx.folder("xl").folder("_rels").file("workbook.xml.rels",relsXML.end({ pretty: true, indent: '  ', newline: '\n' }));

		var buffer = xlsx.generate({type:"nodebuffer"});

		if(response == undefined){
			fs.writeFile(fileName, buffer, function(err) {
			  if (err) throw err;
			});
		}else{
			response.writeHead(200,{
				'Content-Length':buffer.length,
				'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				'Content-Disposition':'attachment; filename='+fileName,
			});
			response.end(buffer);
		}

	}

	return this;
}

exports.WorkBook.prototype.WorkSheet = function(name){
	var newWS = new ws.WorkSheet(name);
	newWS.wb = this;
	this.workbook.WorkSheets.push(newWS);
	return newWS;
}

exports.WorkBook.prototype.Style = function(){
	var newStyle = new style.Style();
	newStyle.wb = this;
	return newStyle;
}

function generateRId(){
    var text = "R";
    var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

    for( var i=0; i < 16; i++ )
        text += possible.charAt(Math.floor(Math.random() * possible.length));

    return text;
}

Number.prototype.toExcelAlpha = function(isCaps){
	//converts column number to text equivalent for Excel
	var isCaps = isCaps == undefined?true:isCaps;

    var d = (this - 1) / 26;
    d = Math.floor(d);
    if (d > 0){
        r = (isCaps ? 65 : 97) + (d - 1);
    }

    if (this % 26 > 0){
        num = (this - (26 * d)) % 26;
    }else{
        num = 26;
    }

    var c = (isCaps ? 65 : 97) + (num - 1);
    if (d > 0){
        return String.fromCharCode(r) + String.fromCharCode(c);
    }else{
        return String.fromCharCode(c);
    }  
}
