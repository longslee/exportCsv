/**
 * 从页面导出Excel的工具类
 * 
 * 这一版实现简单导出，仅支持一个数据源和一个Sheet
 * 
 * @param {JSON} TitleCols  表头行json
 * @param {Object[]} dataSet  数据集数组
 * @param {String} fileName	文件名
 */
var ExportCsv = function(titleCols, dataSet, fileName) {
	this.titleCols = titleCols;
	this.dataSet = dataSet;
	this.fileName = fileName;
}

ExportCsv.prototype.export = function() {
	var titleCols = this.titleCols;
	var dataSet = this.dataSet;
	var fileName = this.fileName;
	saveAs(titleCols, dataSet, fileName);
}

function saveAs(titleCols, dataSet, fileName) {
	if(dataSet.length <= 0){  //没有记录则直接返回
		return;
	}
	var excelFile = constructHTML(titleCols, dataSet, fileName);
	var uri = 'data:application/vnd.ms-excel.sheet.binary.macroEnabled.main;charset=utf-8,' + encodeURIComponent(excelFile);
	var link = document.createElement("a");
	link.href = uri;
	link.download = fileName + ".xls";
	document.body.appendChild(link);
	link.click();
	document.body.removeChild(link);
}

function constructXml(fileName) {
	var xmlStr = "<xml>";
	xmlStr += "<x:ExcelWorkbook>";
	xmlStr += "<x:ExcelWorksheets>";
	xmlStr += "<x:ExcelWorksheet>";
	xmlStr += "<x:Name>";
	xmlStr += fileName;
	xmlStr += "</x:Name>";
	xmlStr += "<x:WorksheetOptions>";
	xmlStr += "<x:DisplayGridlines/>";
	xmlStr += "</x:WorksheetOptions>";
	xmlStr += "</x:ExcelWorksheet>";
	xmlStr += "</x:ExcelWorksheets>";
	xmlStr += "</x:ExcelWorkbook>";
	xmlStr += "</xml>";
	return xmlStr;
}

function constructHTML(titleCols, dataSet, fileName) {
	var excelFile = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>";
	excelFile += '<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">';
	excelFile += "<head>";
	excelFile += constructXml(fileName);
	excelFile += "</head>";
	excelFile += "<body>";
	excelFile += constructTable(titleCols, dataSet);
	excelFile += "</body>";
	excelFile += "</html>";
	return excelFile;
}

function constructTable(titleCols, dataSet) {

	var table = "<table>"
	table += "<tr>";
	$.each(titleCols, function(i, column) { //拼表头
		if(!(column.exprt === undefined) && !column.exprt){
			return true;
		}
		table += "<td>" + column.title + "</td>"
	});
	table += "</tr>";

	$.each(dataSet, function(i, rowData) { //每行数据集
		table += "<tr>";
		$.each(titleCols, function(i, column) { //一行的每列
			if(!(column.exprt === undefined) && !column.exprt){
				return true;
			}
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
			table += "<td style="+column.width+">" + rowContent + "</td>"
		});
		table += "</tr>";
	});
	table += "</table>";

	return table;
}