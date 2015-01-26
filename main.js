var sql = require('mssql'); 
var e4n = require('excel4node');

var config_admin = {
    user: 'adminArek',
    // password: 'KMsa2014',
    password: 'master_2015',
    server: 'ZZM-FILESTR.ZZM.COM.PL\\SQLEXPRESS', // You can use 'localhost\\instance' to connect to named instance
//    database: 'ZZM-FILESTR.ZZM.COM.PL\SQLEXPRESS',

    options: {
        encrypt: true // Use this if you're on Windows Azure
    }
}

var config = {
    user: 'user1',
    password: 'kopex',
    server: 'ZZM-FILESTR.ZZM.COM.PL\\SQLEXPRESS',
    options: {
        encrypt: true // Use this if you're on Windows Azure
    }
}

function pad(num) {
    var s = "00" + num;
    return s.substr(s.length-2);
}

var monthNames = [ "Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
    "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień" ];

/*****************************************************************************/

var connection = new sql.Connection(config);

connection.connect(function(err) {
	if (err) {	// ... error checks
		console.dir(err);
		connection.close();
	} else {
		var request = new sql.Request(connection);
		request.query('select * from PNU.dbo.projekty', function(err, recordset) {
			if (err) {
				// ... error checks
				console.dir(err);
			} else {
				console.dir(recordset[0]);
//				makeXLSX("Excel.xlsx", recordset);
			}
		});
		request.query('select * from PNU.dbo.projekty', function(err, recordset) {
			if (err) {
				// ... error checks
				console.dir(err);
			} else {
				console.dir(recordset[1]);
//				makeXLSX("Excel.xlsx", recordset);
			}
		});
	}
});

/*****************************************************************************/

function firstPage(wb, d) {
	var ws = wb.WorkSheet('1');
	var firstStyle = wb.Style();
	firstStyle.Font.Family('Times New Roman');
	firstStyle.Font.Size(14);
	firstStyle.Font.Alignment.Horizontal('center');
	firstStyle.Font.Color('000000');
	
	ws.Column(4).Width(15);
	ws.Cell(1,1).String('Kopex Machinery S.A.').Style(firstStyle).Format.Font.Size(12).Alignment.Horizontal('left');
	ws.Cell(1,7).String('Zabrze, '+pad(d.getDate())+'.'+pad(d.getMonth()+1)+'.'+d.getFullYear()+' r.').Style(firstStyle).Format.Font.Size(12).Alignment.Horizontal('right');
	ws.Cell(2,1).String('Dział DR i W').Style(firstStyle).Format.Font.Size(12).Alignment.Horizontal('left');
	ws.Cell(5,7).String('ZASTRZEŻONE').Style(firstStyle).Format.Font.Bold().Alignment.Horizontal('right');
	ws.Cell(14,1,14,7,true).String('PLAN  NOWYCH  URUCHOMIEŃ').Style(firstStyle).Format.Font.Size(18).Bold();
	ws.Cell(33,2).String('Opracował').Style(firstStyle);
	ws.Cell(33,6).String('Zatwierdził').Style(firstStyle);
	ws.Cell(36,2).String('. . . . . . . . . . .').Style(firstStyle);
	ws.Cell(36,6).String('. . . . . . . . . . .').Style(firstStyle);
	ws.Cell(38,2).String('Dyrektor DR i W').Style(firstStyle);
	ws.Cell(45,4).String(monthNames[d.getMonth()]+' '+d.getFullYear()+' r. ').Style(firstStyle);
}

function PNUPage(wb, sqlData){
	var PNUtable = wb.Style();
	PNUtable.Font.Color('000000');
	PNUtable.Font.Family('Times New Roman');
	PNUtable.Font.Size(12);
	PNUtable.Border({
		left:{
			style:'thin'
		},
		right:{
			style:'thin'
		},
		top:{
			style:'thin'
		},
		bottom:{
			style:'thin'
		},
		inside:{
			style:'thin'
		}
	});
//	PNUtable.Font.WrapText(true);
	var styleCenter = PNUtable.Clone();
	var styleJustify = PNUtable.Clone();
	styleCenter.Font.Alignment.Horizontal('center');
	styleJustify.Font.Alignment.Horizontal('justify');

/*****************************************************************************/

	var ws = wb.WorkSheet('PNU');
	var columnDefinitions = [
		{head:'L.p.', col:1, style: styleCenter, width: 5},
		{head:'Treść zadania', col:2, style: styleJustify, width: 50},
		{head:'Nr projektu', col:3, style: styleCenter, width: 12},
		{head:'Nr załącznika', col:4, style: styleCenter, width: 13}
	]

	var curRow = 1;
	columnDefinitions.forEach(function(i){
		ws.Cell(curRow,i.col).Style(i.style).String(i.head);
		if (i.col == 1) {
			ws.sheet.cols[0].col['@width'] = i.width;	//!!!!!!!!!
		} else {
			ws.Column(i.col).Width(i.width);
		}
	});
//	console.log(ws.sheet.cols);

	var recordset2 = [];
	sqlData.forEach(function(i){
		if (!i.oddzial) {
			i.oddzial = 'Z';
		} else {
			i.oddzial = i.oddzial.trim();
		}
		if (!recordset2[i.oddzial]) recordset2[i.oddzial] = [];
		recordset2[i.oddzial].push(i);
	});
//	console.dir(recordset2);

	curRow+=1;
	ws.Cell(curRow,1,curRow,4,true).Style(PNUtable).String('Zabrze ścianowe');
	curRow+=1;
	var lp = 1;
	if (recordset2['Z']) {
		recordset2['Z'].forEach(function(i){
			ws.Cell(curRow,columnDefinitions[0].col).Style(columnDefinitions[0].style).String(lp + '.');
			lp += 1;
			ws.Cell(curRow,columnDefinitions[1].col).Style(columnDefinitions[1].style).String(i.opis);
			ws.Cell(curRow,columnDefinitions[2].col).Style(columnDefinitions[2].style).String(i.nr);
			ws.Cell(curRow,columnDefinitions[3].col).Style(columnDefinitions[3].style).String('- zał. nr ' + i.zalacznik);
	//		ws.Row(curRow).Height(15*(Math.floor(i.opis.length/50)+1));	//ceil
			curRow+=1;
		});
	}
	if (recordset2['R']) {
		ws.Cell(curRow,1,curRow,4,true).Style(PNUtable).String('Rybnik');
		curRow+=1;
		recordset2['R'].forEach(function(i){
			ws.Cell(curRow,columnDefinitions[0].col).Style(columnDefinitions[0].style).String(lp + '.');
			lp += 1;
			ws.Cell(curRow,columnDefinitions[1].col).Style(columnDefinitions[1].style).String(i.opis);
			ws.Cell(curRow,columnDefinitions[2].col).Style(columnDefinitions[2].style).String(i.nr);
			ws.Cell(curRow,columnDefinitions[3].col).Style(columnDefinitions[3].style).String('- zał. nr ' + i.zalacznik);
	//		ws.Row(curRow).Height(15*(Math.floor(i.opis.length/50)+1));	//ceil
			curRow+=1;
		});
	}
	if (recordset2['W']) {
		ws.Cell(curRow,1,curRow,4,true).Style(PNUtable).String('Zabrze chodnikowe');
		curRow+=1;
		recordset2['W'].forEach(function(i){
			ws.Cell(curRow,columnDefinitions[0].col).Style(columnDefinitions[0].style).String(lp + '.');
			lp += 1;
			ws.Cell(curRow,columnDefinitions[1].col).Style(columnDefinitions[1].style).String(i.opis);
			ws.Cell(curRow,columnDefinitions[2].col).Style(columnDefinitions[2].style).String(i.nr);
			ws.Cell(curRow,columnDefinitions[3].col).Style(columnDefinitions[3].style).String('- zał. nr ' + i.zalacznik);
	//		ws.Row(curRow).Height(15*(Math.floor(i.opis.length/50)+1));	//ceil
			curRow+=1;
		});
	}
	if (recordset2['D']) {
		ws.Cell(curRow,1,curRow,4,true).Style(PNUtable).String('Zabrze DHTP');
		curRow+=1;
		recordset2['D'].forEach(function(i){
			ws.Cell(curRow,columnDefinitions[0].col).Style(columnDefinitions[0].style).String(lp + '.');
			lp += 1;
			ws.Cell(curRow,columnDefinitions[1].col).Style(columnDefinitions[1].style).String(i.opis);
			ws.Cell(curRow,columnDefinitions[2].col).Style(columnDefinitions[2].style).String(i.nr);
			ws.Cell(curRow,columnDefinitions[3].col).Style(columnDefinitions[3].style).String('- zał. nr ' + i.zalacznik);
	//		ws.Row(curRow).Height(15*(Math.floor(i.opis.length/50)+1));	//ceil
			curRow+=1;
		});
	}
}

function x(request, sql_query, callback){
	request.query(sql_query, function(err, recordset) {
		// ... error checks
		if (err) {
			console.dir(err);
		} else {
//			console.dir(recordset[0]);
			callback(err, recordset);
		}
	});
	
}

function makeXLSX(filename, request){
	var wb = new e4n.WorkBook();
	firstPage(wb, new Date());
	x(request, 'select * from PNU.dbo.projekty', function(err, sql_data){
		if (err) return console.error(err);
		PNUPage(wb, sql_data);
	});
		
	wb.write(filename);	// Synchronously write file
	connection.close();
}
