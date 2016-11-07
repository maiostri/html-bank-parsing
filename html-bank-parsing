var jsdom = require("jsdom");
var fs = require('fs');
var wgxpath = require('wgxpath');
var excelbuilder = require('msexcel-builder');
var expressionString = "//table[@class='TRNfundo']//td[text()='Lançamentos nacionais']";

jsdom.env({
	file: process.argv[2],
	done: function (err, window) {
		GLOBAL.window = window;
		GLOBAL.document = window.document;
		wgxpath.install(window);
		var expression = window.document.createExpression(expressionString);
		var result = expression.evaluate(window.document, wgxpath.XPathResultType.ANY_TYPE).iterateNext().parentNode.parentNode.children;
		parseFile(result);
	}
});


function parseElements(elements) {

	var lanc = [];

	for (var i = 0; i < elements.length; i++) {

		var individualLanc = {};

		var children = elements[i].children;
		var firstField = false;
		for (var j=0; j < children.length; j++) {

			if (children[j].textContent.includes('Crédito') || children[j].textContent.includes('Débito')) {
				continue;
			}

			if (children[j].hasAttribute('width')) {
				continue;
			}

			if (children[j].className == 'TRNcampo_linha') {
				if (firstField) {
					individualLanc.text = children[j].textContent.trimRight();
				}
				else {
					date = children[j].textContent.trim().replace(/\//g,'-');
					date = date + '-2016';
					individualLanc.date = date.toString();	

					firstField = true;
				}
			}
			else if (children[j].className == 'TRNtitcampo_linha') {
				var value = children[j].textContent.trim().replace(/,/g,'.');
				console.log(value);
				individualLanc.value = (parseFloat(value) * (-1)).toString();
			}
		}

		if (individualLanc && individualLanc.text) {
			lanc.push(individualLanc);
		}
	}

	return lanc;
}

function createFile(elements) {
	var options = { encoding: 'utf8/utf-8' };
	var workbook = excelbuilder.createWorkbook('./', 'sample.xlsx')
	var sheet1 = workbook.createSheet('sheet1', 10, 100);

	sheet1.set(1, 1, 'Data');
	sheet1.set(2, 1, 'Descrição');
	sheet1.set(3, 1, 'Categoria');
	sheet1.set(4, 1, 'Valor');

	var line = 4;
	elements.forEach(function (element) {
		sheet1.set(1, line, element.date);
		sheet1.set(2, line, element.text);
		sheet1.set(4, line, element.value.toString().replace(/\./g,','));
		line++;
	});



// Save it
  workbook.save(function(err){
    if (err)
      throw err;
    else
      console.log('congratulations, your workbook created');
  });

	// var writeStream = fs.createWriteStream("file.xls");
	// var header="Data"+"\t"+"Descrição"+"\t"+"Categoria"+"\t"+"Valor"+"\n";
	// writeStream.write(header);
	// elements.forEach(function (element) {
	// 	var row = element.date+"\t"+element.text+"\t"+"\t"+element.value.toString().replace(/\./g,',')+"\n";
	// 	console.log(row);
	// 	writeStream.write(row);
	// });

	// writeStream.end();
}

function parseFile(elements) {
	var result = parseElements(elements);
	console.log(result);
	createFile(result);
}