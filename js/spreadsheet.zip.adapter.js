/*jshint -W083 */
( function SpreadsheetZipAdapter(URL,sheet,dataCb,progressCb){

	function getParsedText(text){
		var xmlDoc, ActiveXObject;
		if (window.DOMParser){
			var parser;
			parser = new DOMParser();
			xmlDoc = parser.parseFromString(text,"text/xml");
		}else{
			//IE
			xmlDoc =new ActiveXObject("Microsoft.XMLDOM");
			xmlDoc.async = false;
			xmlDoc.loadXML(text); 
		}
		return xmlDoc;
	}

	function readSharedStrings(text){
		var xmlDoc = getParsedText(text);
		var sis = xmlDoc.getElementsByTagName('si');
		var result = {};
		for (var i = 0; i < sis.length; i++){
			result[i] = sis[i].getElementsByTagName('t')[0].childNodes[0].nodeValue;
		}
		return result;
	}

	function readData(sst,text,progressCb){
		var data = {};
		var xmlDoc = getParsedText(text);
		var cells = xmlDoc.getElementsByTagName('c');
		var len = cells.length;
		var mod = 1;
		if (len >= 100){
			mod = Math.floor(len/100);
		}
		window.console.log(len);
		window.console.log(mod);
		for(var c = 0; c < len; c++){
			(function(cell){
				var type = cell.getAttribute('t');
				var pos = cell.getAttribute('r');
				var formula;
				//console.log(pos);
				//console.log(cell);
				switch(type){
					case null:
						formula = getValue(cell);
						break;
					case 'e':
					case 'n':
					case 'str':
						formula = getFormulaOrValue(sst,cell);
						break;
					case 's':
						formula = getFormulaOrValue(sst,cell);
						if(formula){
							if(!(formula in sst)) throw "could not find value in sst";
							formula = sst[formula];
						}
						break;
					default:
						throw "unhandled type "+ type;
				}
				data[pos] = {
					formula: formula
				};
				if(c % mod === 0){
					progressCb(c,len);
				}
			})(cells[c]);
		}
		progressCb(c,len);
		return data;		
	}

	function getFormula(sst,cell){
		var elements = cell.getElementsByTagName('f');
		if(elements.length){
			var child = elements[0];
			var type = child.getAttribute('t');
			if(type && type === 'shared'){
				var si = child.getAttribute('si');
				return sst[si];
			}
			return '='+getValueOfElement(child);
		}
	}

	function getValue(cell){
		var elements = cell.getElementsByTagName('v');
		if(elements.length){
			return getValueOfElement(elements[0]);
		}
		return;
	}

	function getFormulaOrValue(sst,cell){
		var result = getFormula(sst,cell);
		if(!result){
			result = getValue(cell);
		}
		return result;
	}

	function getValueOfElement(e){
		if(e.childNodes && e.childNodes.length){
			return e.childNodes[0].nodeValue;
		}
		return;
	}

	function getEntry(filename,callback) {
		this.getEntries(function(entries){
			for(var x = 0; x < entries.length; x++){
				if(entries[x].filename === filename){
					callback(entries[x]);		
				}
			}
		});
	}

	var zip;

	if(!zip){
		throw "Requires zip library";
	}
	if( window.Worker ){
		//TODO this should probably not be hardcoded like this
		zip.workerScriptsPath = "/extensions/Spreadsheet/lib/zip/WebContent/";
	}else{
		zip.useWebWorkers = false;
	}

	zip.createReader(new zip.HttpReader(URL), function(reader) {
		window.console.log(reader);
		reader.getEntry = getEntry;

		reader.getEntry('xl/sharedStrings.xml',function(entry){
			entry.getData(new zip.TextWriter(), function(text) {
				var sst = readSharedStrings(text);

				var sheetFileName = 'xl/worksheets/sheet'+(sheet+1)+'.xml';
				window.console.log(sheetFileName);
				var sheetEntry = reader.getEntry(sheetFileName,function(entry){
					entry.getData(new zip.TextWriter(), function(text){
						dataCb(readData(sst,text,progressCb));
					},progressCb);
				});

			});
		});
		
		reader.close();
	});/*,*/
	/*
	function(error) {
	// onerror callback
	}
	*/
	/*);*/
}());