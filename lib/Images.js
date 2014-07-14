var fs = require('fs'),
mime = require('mime'),
imgsz = require('image-size');

var drawing = function(imgURI){
	var d = {
		props:{
			image:imgURI,
			imageId:0,
			mimeType:'',
			extension:'',
			width:0,
			height:0,
			dpi:96
		},
		xml:{			
			'xdr:oneCellAnchor':[
				{
					'xdr:from':{
						'xdr:col':0,
						'xdr:colOff':0,
						'xdr:row':0,
						'xdr:rowOff':0
					}
				},
				{
					'xdr:ext':{
						'@cx':0*9525,
						'@cy':0*9525
					}
				},
				{
					'xdr:pic':{
						'xdr:nvPicPr':[
							{
								'xdr:cNvPr': {
									'@descr':'image',
									'@id':0,
									'@name':'Picture'
								}
							},
							"xdr:cNvPicPr"
						],
						'xdr:blipFill':{
							'a:blip':{
								'@cstate':'print',
								'@r:embed':'rId0',
								'@xmlns:r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
							}
						},
						'xdr:spPr':[
							{
								'@bwMode':'auto'
							},
							{
								'a:xfrm':{
									'a:off':{
										'@x':0,
										'@y':0
									},
									'a:ext':{
										'@cx':0*9525,
										'@cy':0*9525
									}
								}
							},
							{
								'a:prstGeom':[
									{
										'@prst':'rect'
									},
									'a:avLst'
								]
							},
							"a:noFill",
							{
								'a:ln':['a:noFill']
							}
						]
					}
				},
				"xdr:clientData"
			] 
		},
		Position:function(r, c, offY, offX){
			var offsetX = offX?offX:0;
			var offsetY = offY?offY:0;

			d.xml['xdr:oneCellAnchor'].forEach(function(v){
				if(v['xdr:from']){
					v['xdr:from']['xdr:col']=r-1;
					v['xdr:from']['xdr:row']=c-1;
					v['xdr:from']['xdr:colOff']=offsetX;
					v['xdr:from']['xdr:rowOff']=offsetY;
				}
				else if(v['xdr:pic'] && offX && offY){
					var spPR = v['xdr:pic']['xdr:spPr'];
					spPR.forEach(function(o){
						if(o['a:xfrm']){
							o['a:xfrm']['a:off']['@x']=offsetX;
							o['a:xfrm']['a:off']['@y']=offsetY;
						}
					});
				}
			});
		},
		Properties:function(props){
			Object.keys(props).forEach(function(k){
				d.props[k] = props[k];
			});
		},
		SetID:function(id){
			d.xml['xdr:oneCellAnchor'].forEach(function(v){
				if(v['xdr:pic']){
					v['xdr:pic']['xdr:nvPicPr'][0]['xdr:cNvPr']['@id'] = id;
					v['xdr:pic']['xdr:blipFill']['a:blip']['@r:embed'] = 'rId'+id;
				}
			});
			d.props.imageId=id;
		},
		updateSize:function(){
			d.xml['xdr:oneCellAnchor'].forEach(function(v){
				if(v['xdr:ext']){
					v['xdr:ext']['@cx']=d.props.width*9525*(96/d.props.dpi);
					v['xdr:ext']['@cy']=d.props.height*9525*(96/d.props.dpi);
				}
				else if(v['xdr:pic']){
					var spPR = v['xdr:pic']['xdr:spPr'];
					spPR.forEach(function(o){
						if(o['a:xfrm']){
							o['a:xfrm']['a:ext']['@cx']=d.props.width*9525*(96/d.props.dpi);
							o['a:xfrm']['a:ext']['@cy']=d.props.height*9525*(96/d.props.dpi);
						}
					});
				}
			});
		}

	}

	return d;
}

exports.Image = function(imgURI){

	var wb=this.wb.workbook;
	var ws=this;

	// add entry to [Content_Types].xml
	var mimeType = mime.lookup(imgURI);
	var extension = mimeType.split('/')[1];

	var contentTypeAdded = false;
	wb.Content_Types.Types.forEach(function(t){
		if(t['Default']){
			if(t['Default']['@ContentType'] == mimeType){
				contentTypeAdded = true;
			}
		}
	})
	if(!contentTypeAdded){
		wb.Content_Types.Types.push({
			"Default":{
				"@ContentType":mimeType,
				"@Extension":extension
			}
		});
	}

	// create drawingn.xml file
	// create drawingn.xml.rels file
	if(!ws.drawings){
		ws.drawings = {
			'rels':{
				'Relationships':[
					{
						'@xmlns':'http://schemas.openxmlformats.org/package/2006/relationships'
					}
				]
			},
			'xml':{
				'xdr:wsDr':[
					{
						'@xmlns:a':'http://schemas.openxmlformats.org/drawingml/2006/main',
						'@xmlns:xdr':'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
					}
				]
			},
			drawings:[]
		};
	}
	if(!ws.rels){
		ws.rels = {
			'Relationships':[
				{
					'@xmlns':'http://schemas.openxmlformats.org/package/2006/relationships'
				}
			]
		}
	}

	var d = new drawing(imgURI);


	d.Properties({
		'mimeType':mimeType,
		'extension':extension
	});

	var dim = imgsz(imgURI);
	d.Properties({
		'width':dim.width,
		'height':dim.height
	});
	d.updateSize();

	ws.drawings.drawings.push(d);
	var imgID = 0;
	wb.WorkSheets.forEach(function(s){
		if(s.drawings){
			imgID+=s.drawings.drawings.length;
		}
	});
	d.SetID(imgID);

	ws.drawings.rels.Relationships.push({
		'Relationship':{
			'@Id':'rId'+imgID,
			'@Target':'../media/image'+imgID+'.'+extension,
			'@Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
		}
	});

	var relExists = false;
	ws.rels['Relationships'].forEach(function(r){
		if(r['Relationship']){
			if(r['Relationship']['@Id'] == 'rId1'){
				relExists = true;
			}
		}
	});
	if(!relExists){
		ws.rels['Relationships'].push({
			'Relationship':{
				'@Id':'rId1',
				'@Target':'../drawings/drawing'+ws.sheetId+'.xml',
				'@Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
			}
		});
	}
	ws.sheet.drawing = {
		'@r:id':'rId1'
	}

	return d;
}