/**
 * 从页面导出Excel的工具类
 * 
 * 这一版实现了多数据源和多Sheets
 * 
 * @param {Object[JSON[]]} titleCols  表头行
 * @param {String[]} wsnames  sheets name
 * @param {Object[Object[]]} dataSets  数据集数组
 * @param {String} fileName	文件名
 */

var ExportCsv = function(titleCols, wsnames, dataSets, fileName) {
	this.titleCols = titleCols;
	this.wsnames = wsnames;
	this.dataSets = dataSets;
	this.fileName = fileName;
}

ExportCsv.prototype.export = function() {
	var titleCols = this.titleCols;
	var wsnames = this.wsnames;
	var dataSets = this.dataSets;
	var fileName = this.fileName;
	saveAs(titleCols, wsnames, dataSets, fileName);
}

function saveAs(titleCols, wsnames, dataSets, fileName) {
	if(dataSets.length <= 0) { //没有记录则直接返回
		return;
	}
	var uri = 'data:application/vnd.ms-excel;base64,';
	var tmplWorkbookXML = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">' +
		'<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office"><Author>Staryea penglongjia</Author><Created>{created}</Created></DocumentProperties>' +
		'{worksheets}</Workbook>';
	var tmplWorksheetXML = '<Worksheet ss:Name="{nameWS}"><Table>{rows}</Table></Worksheet>';
	var tmplCellXML = '<Cell><Data ss:Type="{nameType}">{data}</Data></Cell>';
	var workbookXML = constructWbXML(titleCols, wsnames, dataSets, fileName, tmplWorkbookXML, tmplWorksheetXML, tmplCellXML);
	var link = document.createElement("A");
	link.href = uri + base64(workbookXML);
	link.download = fileName+'.xls' || 'Workbook.xls';
	link.target = '_blank';
	document.body.appendChild(link);
	link.click();
	document.body.removeChild(link);
}

function constructWbXML(titleCols, wsnames, dataSets, fileName, tmplWorkbookXML, tmplWorksheetXML, tmplCellXML) {
	var ctx = "";
	var workbookXML = "";
	var worksheetsXML = "";
	var rowsXML = "";

	$.each(titleCols, function(i, columns) { //多个表头

		//拼表头 start
		rowsXML += '<Row>';
		$.each(columns, function(j, column) {
			var ctx = {
				nameType: 'String',
				data: column.title
			};
			rowsXML += format(tmplCellXML, ctx);
		});
		rowsXML += '</Row>';
		//拼表头end

		//拼表内容  
		//因为dataSets与titleCols下表对应，所以直接区
		$.each(dataSets[i], function(k, rowData) {
			rowsXML += '<Row>';
			$.each(columns, function(j, column) { //一行的每列
				var data = column.data;
				var width = column.width;
				var defaultContent = column.defaultContent;
				//如果是多层节点  需要处理
				var rowContent;
				var splitData = data.split('.');
				if(splitData.length > 1) {
					rowContent = rowData[splitData[0]]; //第一个
					for(var i = 1; i < splitData.length; i++) {
						rowContent = rowContent[splitData[i]];
					}
				} else {
					rowContent = rowData[data];
				}
				if(rowContent === undefined) rowContent = defaultContent;
				var dealCallback = column.dealCallback;
				if(!(dealCallback === undefined)) {
					rowContent = dealCallback(rowContent);
				}
				var ctx = {
					nameType: 'String',
					data: rowContent
				};
				rowsXML += format(tmplCellXML, ctx);
			});
			rowsXML += '</Row>';
		});

		var ctx = {
			rows: rowsXML,
			nameWS: wsnames[i] || 'Sheet' + i
		};
		worksheetsXML += format(tmplWorksheetXML, ctx);
		rowsXML = "";
	});
	var ctx = {
		created: (new Date()).getTime(),
		worksheets: worksheetsXML
	};
	workbookXML = format(tmplWorkbookXML, ctx);
	return workbookXML;
}

function base64(s) {
	return window.btoa(unescape(encodeURIComponent(s)));
}

function format(s, c) {
	return s.replace(/{(\w+)}/g, function(m, p) {
		return c[p];
	});
}