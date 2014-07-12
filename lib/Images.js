var fs = require('fs'),
mime = require('mime');

var drawing = function(imgURI){

	var d = {
		image:imgURI,
		xml:{
			'xdr:wsDr':{
				'@xmlns:a':'http://schemas.openxmlformats.org/drawingml/2006/main',
				'@xmlns:xdr':'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
				'xdr:twoCellAnchor':{
					'@editAs':'oneCell',
					'xdr:from':{
						'xdr:col':0
					},
					'xdr:to':{
						'xdr:col':1
					},
					'xdr:pic':{
						'xdr:nvPicPr':{
							'xdr:cNvPr':{
								'@descr':'uri'
							}
						}
					}
				} 
			}
		},
		Size:function(w,h){
			console.log([w,h]);
			console.log(d);
		},
		Position:function(r,c){
			console.log([r,c]);
			console.log(d);
		}

	}

	return d;
}

exports.Image = function(imgURI){

	var wb=this.wb.workbook;
	var ws=this.sheet;


	// add entry to [Content_Types].xml
	var mimeType = mime.lookup(imgURI);
	var extension = mimeType.split('/')[1];
	wb.Content_Types.Types.push({
		"Default":{
			"@ContentType":mimeType,
			"@Extension":extension
		}
	})


	// create drawingn.xml file
	// create drawingn.xml.rels file
	if(!wb.drawings){
		wb.drawings = {
			'rels':{
				'Relationships':[
					{
						'@xmlns':'http://schemas.openxmlformats.org/package/2006/relationships'
					}
				]
			},
			'drawings':[]
		};
	}

	var d = new drawing(imgURI);
	var imgID = wb.drawings.drawings.push(d);

	wb.drawings.rels.Relationships.push({
		'Relationship':{
			'@Id':'rId'+imgID,
			'@Target':'../media/image'+imgID+'.'+extension,
			'@Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
		}
	})

	// add "drawing" element to sheetn.xml
	// add entry to sheetn.xml.rels


	console.log(JSON.stringify(wb.drawings,null,'\t'));
	return d;
}