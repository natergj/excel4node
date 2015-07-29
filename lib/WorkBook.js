var WorkSheet = require('./WorkSheet.js'),
style = require('./Style.js'),
xml = require('xmlbuilder'),
jszip = require('jszip'),
xml = require('xmlbuilder');
fs = require('fs');

var xmlOutVars = {};
var xmlDebugVars = { pretty: true, indent: '  ',newline: '\n' };

var WorkBook = function(){
	var opts = opts?opts:{};

	this.opts = {};
	this.opts.jszip = {};
	this.opts.jszip.compression = 'DEFLATE';
	if(opts.jszip){
		Object.keys(opts.jszip).forEach(function(k){
			this.opts.jszip[k] = opts.jszip.compression;
		});
	};

	this.defaults={
		colWidth:opts.colWidth?opts.colWidth:15
	};
	this.styleData={
		numFmts:[],
		fonts:[],
		fills:[],
		borders:[],
		cellXfs:[]
	};
	this.worksheets=[];
	this.workbook = {
		WorkSheets:[],
		workbook:{
			'@xmlns:r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
			'@xmlns':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
			bookViews:[
				{
					workbookView:{
						'@tabRatio':'600',
						'@windowHeight':'14980',
						'@windowWidth':'25600',
						'@xWindow':'0',
						'@yWindow':'1080'
					}
				}
			],
			sheets:[]
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
		debug:false
	}

	// Generate default style
	this.Style();

	return this
}

WorkBook.prototype.WorkSheet = function(name, opts){
	var thisWS = new WorkSheet(this);
	thisWS.setName(name);
	thisWS.setWSOpts(opts);
	thisWS.sheetId = this.worksheets.length + 1;
	this.worksheets.push(thisWS);
	return thisWS;
}

WorkBook.prototype.getStringIndex = function(val){

	if(this.workbook.sharedStrings.indexOf(val) < 0){
		this.workbook.sharedStrings.push(val)
	};
	return this.workbook.sharedStrings.indexOf(val);
}

WorkBook.prototype.getStringFromIndex = function(val){

	return this.workbook.sharedStrings[val];
}

WorkBook.prototype.write = function(fileName, response){

	var buffer = this.writeToBuffer();

	// If `response` is an object (a node response object)
	if(typeof response === "object"){
		response.writeHead(200,{
			'Content-Length':buffer.length,
			'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			'Content-Disposition':'attachment; filename='+fileName,
		});
		response.end(buffer);
	}

	// Else if `response` is a function, use it as a callback
	else if(typeof response === "function"){
		fs.writeFile(fileName, buffer, function(err) {
			response(err);
		});
	}

	// Else response wasn't specified
	else {
		fs.writeFile(fileName, buffer, function(err) {
			if (err) throw err;
		});
	}
}

WorkBook.prototype.createStyleSheetXML = function(){
	var thisWB = this;
	var data={
			styleSheet:{
				'@xmlns':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
				'@mc:Ignorable':'x14ac',
				'@xmlns:mc':'http://schemas.openxmlformats.org/markup-compatibility/2006',
				'@xmlns:x14ac':'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
				numFmts:[],
				fonts:[],
				fills:[],
				borders:[],
				cellXfs:[]
			}
		}


	var items = [
		'numFmts',
		'fonts',
		'fills',
		'borders',
		'cellXfs'
	];

	items.forEach(function(i){
		data.styleSheet[i].push({'@count':thisWB.styleData[i].length});
		thisWB.styleData[i].forEach(function(d){
			data.styleSheet[i].push(d.generateXMLObj());
		});
	});

	var styleXML = xml.create(data);
	return styleXML.end(xmlOutVars);
}

WorkBook.prototype.writeToBuffer = function(){
	var xlsx = new jszip();
	var that = this;
	var wbOut = JSON.parse(JSON.stringify(this.workbook));

	this.worksheets.forEach(function(sheet, i){
		var sheetCount = i+1;
		var thisRId = generateRId();
		var sheetExists = false;

		wbOut.workbook.sheets.forEach(function(s){
			if(s.sheet['@sheetId'] == sheetCount){
				sheetExists = true;
			}
		});
		if(!sheetExists){
			wbOut.workbook_xml_rels.Relationships.push({
				Relationship:{
					'@Id':thisRId,
					'@Target':'worksheets/sheet'+sheetCount+'.xml',
					'@Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
				}
			})
			wbOut.workbook.sheets.push({
				sheet:{
					'@name':sheet.name,
					'@sheetId':sheetCount,
					'@r:id':thisRId
				}
			});
			wbOut.Content_Types.Types.push({
				Override:{
					'@ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
					'@PartName':'/xl/worksheets/sheet'+sheetCount+'.xml'
				}
			});
			if(that.debug){
				console.log("\n\r###### Sheet XML XML #####\n\r");
				//console.log(xmlStr.end(xmlDebugVars))
			};
		}

		if(sheet.drawings){
			if(that.debug){
				console.log("\n\r########  Drawings found ########\n\r")
			}
			var drawingRelsXML = xml.create(sheet.drawings.rels);
			if(that.debug){
				console.log("\n\r###### Drawings Rels XML #####\n\r");
				console.log(drawingRelsXML.end(xmlDebugVars))
			};
			xlsx.folder("xl").folder("drawings").folder("_rels").file("drawing"+sheet.sheetId+".xml.rels",drawingRelsXML.end(xmlOutVars));

			sheet.drawings.drawings.forEach(function(d){
				sheet.drawings.xml['xdr:wsDr'].push(d.xml);
				xlsx.folder("xl").folder("media").file('image'+d.props.imageId+'.'+d.props.extension,fs.readFileSync(d.props.image));
				if(that.debug){
					console.log("\n\r###### Drawing image data #####\n\r");
					console.log(fs.statSync(d.props.image))
				};
			});

			var drawingXML = xml.create(sheet.drawings.xml);
			xlsx.folder("xl").folder("drawings").file("drawing"+sheet.sheetId+".xml",drawingXML.end(xmlOutVars));
			if(that.debug){
				console.log("\n\r###### Drawings XML #####\n\r");
				console.log(drawingXML.end(xmlDebugVars))
			};

			wbOut.Content_Types.Types.push({
				Override:{
					'@ContentType': 'application/vnd.openxmlformats-officedocument.drawing+xml',
					'@PartName':'/xl/drawings/drawing'+sheet.sheetId+'.xml'
				}
			});
		}

		if(sheet.hyperlinks) {
			wbOut.Content_Types.Types.push({
				Override:{
					'@ContentType': 'application/vnd.openxmlformats-package.relationships+xml',
					'@PartName':'/xl/worksheets/_rels/sheet'+sheet.sheetId+'.xml.rels'
				}
			});
			if(!sheet.rels){
				sheet.rels = {
					'Relationships':[
						{
							'@xmlns':'http://schemas.openxmlformats.org/package/2006/relationships'
						}
					]
				}
			}
			sheet.hyperlinks.forEach(function(cr, i){
				var found = false;
				sheet.rels.Relationships.forEach(function(r){
					if(r.Relationship && r.Relationship['@Id'] == 'rId' + cr.id){
						found = true;
					}
				});

				if(!found){
					sheet.rels.Relationships.push({
						'Relationship':{
							'@Id': 'rId' + cr.id,
							'@Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
							'@Target': cr.url,
							'@TargetMode': 'External'
						}
					});
				}
			});
		}

		if(sheet.rels){
			var sheetRelsXML = xml.create(sheet.rels);
			if(that.debug){console.log(sheetRelsXML.end(xmlDebugVars))};
			xlsx.folder("xl").folder("worksheets").folder("_rels").file("sheet"+sheet.sheetId+".xml.rels",sheetRelsXML.end(xmlOutVars));
		}

		if(sheet.sheet.autoFilter){
			if(!wbOut.workbook['definedNames']){
				wbOut.workbook['definedNames'] = [];
			}
			wbOut.workbook['definedNames'].push({
				definedName : {
					'@hidden' : 1,
					'@localSheetId' : i,
					'@name' : '_xlnm._FilterDatabase',
					'#text': '\'' + sheet.name + '\'!$A$1:$C$1'
				}
			});
		}

		var xmlStr = sheet.toXML();
		xlsx.folder("xl").folder("worksheets").file('sheet'+sheetCount+'.xml',xmlStr);

		if(that.debug){
			console.log("\n\r###### SHEET "+sheetCount+" XML #####\n\r");
			console.log(xmlStr)
		};
	});

	wbOut.sharedStrings.forEach(function(s){
		wbOut.strings.sst.push({'si':{'t':s}});
	});

	wbOut.strings.sst[0]['@uniqueCount']=wbOut.sharedStrings.length;

	var wbXML = xml.create({workbook:JSON.parse(JSON.stringify(wbOut.workbook))});
	if(this.debug){
		console.log("\n\r###### WorkBook XML #####\n\r");
		console.log(wbXML.end(xmlDebugVars));
	};

	//var styleXML = xml.create(JSON.parse(JSON.stringify(wbOut.styles)));
	var styleXMLStr = this.createStyleSheetXML();
	//console.log(styleXMLStr);

	if(this.debug){
		console.log("\n\r###### Style XML #####\n\r");
		console.log(styleXMLStr);
	};

	var relsXML = xml.create(wbOut.workbook_xml_rels);
	if(this.debug){
		console.log("\n\r###### WorkBook Rels XML #####\n\r");
		console.log(relsXML.end(xmlDebugVars))
	};

	var Content_TypesXML = xml.create(wbOut.Content_Types);
	if(this.debug){
		console.log("\n\r###### Content Types XML #####\n\r");
		console.log(Content_TypesXML.end(xmlDebugVars))
	};

	var globalRelsXML = xml.create(wbOut.global_rels);
	if(this.debug){
		console.log("\n\r###### Globals Rels XML #####\n\r");
		console.log(globalRelsXML.end(xmlDebugVars))
	};

	var stringsXML = xml.create(wbOut.strings);
	if(this.debug){
		console.log("\n\r###### Shared Strings XML #####\n\r");
		console.log(stringsXML.end(xmlDebugVars))
	};

	xlsx.file("[Content_Types].xml",Content_TypesXML.end(xmlOutVars));
	xlsx.folder("_rels").file(".rels",globalRelsXML.end(xmlOutVars));
	xlsx.folder("xl").file("sharedStrings.xml",stringsXML.end(xmlOutVars));
	xlsx.folder("xl").file("styles.xml",styleXMLStr);
	xlsx.folder("xl").file("workbook.xml",wbXML.end(xmlOutVars));
	xlsx.folder("xl").folder("_rels").file("workbook.xml.rels",relsXML.end(xmlOutVars));

	delete wbOut;
	return xlsx.generate({type:"nodebuffer",compression:this.opts.jszip.compression});
}

WorkBook.prototype.Style = style.Style;

exports.WorkBook = WorkBook;


function generateRId(){
    var text = "R";
    var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

    for( var i=0; i < 16; i++ )
        text += possible.charAt(Math.floor(Math.random() * possible.length));

    return text;
}