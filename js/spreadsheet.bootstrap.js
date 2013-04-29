$(document).ready(function(){
	function createSpreadsheet(container,cellData){
		return new Spreadsheet.Slick(container,cellData);
	}

	//TODO create conversion logic that supports all possible excel graphs?
	function convertChartDataToHighCharts(c,cellData){
		var hc = {
			start: c.start.cell,
			end: c.end.cell,
			title: c.title  ? c.title[0].text : '',
			series: [],
			xAxis: {
				labels: {}
			},
			chart: {},
			plotOptions: {}
		};

		var plotGroup = c.data.plotGroup[0];
		var data = plotGroup.data;
		switch(plotGroup.type){
			case 'areaChart':
			hc.chart.type = 'area';
			break;
			case 'barChart':
			switch(plotGroup.direction){
				case 'col':
				hc.chart.type = 'column';
				break;
				case 'bar':
				hc.chart.type = 'bar';
				break;
				default:
				throw "unhandled plotGroup direction";
			}
			break;
			case 'scatterChart':
			hc.chart = $.extend(hc.chart,{
				type: 'scatter',
				zoomType: 'xy',
				inverted: 'true'
			});
			hc.xAxis.labels.rotation = 0;

			for(var x = 0; x < data.length; x++){
				if(!hc.series[x]){
					hc.series[x] = {};
				}
				if(!hc.series[x].name){
					hc.series[x].name = data[x].label.dataSource;
				}
				hc.series[x].data = data[x].category.dataSource;
				if(data[x].value && !hc.xAxis.categories){
					hc.xAxis.categories = data[x].value.dataSource;
				}
			}
			break;
			default:
			throw "unhandled chart type";
		}


		var plotOptions = hc.plotOptions[hc.chart.type] = {};
		switch(plotGroup.grouping){
			case 'percentStacked':
			plotOptions.stacking = 'percent';
			break;
			case 'stacked':
			plotOptions.stacking = 'normal';
			break;
			case 'clustered':
			case 'standard':
			case null:
			case undefined:
			break;
			default:
			throw "unhandled chart stacking";
		}

		if(hc.series.length === 0){
			for(var x = 0; x < data.length; x++){
				if(!hc.series[x]){
					hc.series[x] = {};
				}
				if(!hc.series[x].name){
					hc.series[x].name = data[x].label.dataSource;
				}
				hc.series[x].data = data[x].value.dataSource;
				if(data[x].category && !hc.xAxis.categories){
					hc.xAxis.categories = data[x].category.dataSource;
				}
			}
		}

		return hc;
	}

	function getOnLoadCallback(spreadsheet){
		return function(c){
			//on load
			//Add series value change listeners
			for (var i = 0; i < c.series.length; i++) {
				var series = c.series[i];
				for (var j = 0;  j < series.data.length; j++) {
					var point = series.data[j];
					//add listener for point value
					spreadsheet.addValueChangeListener(point.id,getValueChangeListener(c));
				}
			}
		};
	}

	function getValueChangeListener(c){
		return function(e){
			var point = c.get(e.cell.position);
			try{
				point.update(e.cell.valueOf());
			}catch(err){
				alert('Error occured while updating point!');
			}
		};
	}

	function renderChart(config,spreadsheetSlick){
		var grid = spreadsheetSlick.getGrid();
		var spreadsheet = spreadsheetSlick.getSpreadsheet();

		var appendChartInterval = setInterval(appendChart,500);

		function appendChart(){
			var startPos = Spreadsheet.parsePosition(config.start);
			var endPos = Spreadsheet.parsePosition(config.end);

			var startCellNode = grid.getCellNode(startPos.row-1,startPos.getColumnIndex()+1);
			var endCellNode = grid.getCellNode(endPos.row-1,endPos.getColumnIndex()+1);

			if(!(startCellNode && endCellNode)){
				return;
			}
			delete config.start;
			delete config.end;

			var chartStartPos = {
				left: $(startCellNode).position().left,
				top: $(startCellNode).parent().position().top
			};

			var chartEndPos = {
				left: $(endCellNode).position().left,
				top: $(endCellNode).parent().position().top
			};

			var chartX = chartStartPos.top;
			var chartY = chartStartPos.left;
			var chartWidth = chartEndPos.left-chartStartPos.left;
			var chartHeight = chartEndPos.top-chartStartPos.top;

			var chartId = 'chart-'+Date.now();
			var chartNode = $('<div></div>').css({
				position: 'absolute',
				top:chartX,
				left:chartY,
				zIndex:100
			}).attr('id',chartId).appendTo('.grid-canvas');


			config = $.extend(true,{
				chart: {
					renderTo: chartId,
					width: chartWidth,
					height: chartHeight
				},
				xAxis: {
					labels:{
						rotation: -90,
						step: 3,
						style: {
							fontSize: '10px',
							fontFamily: 'Verdana, sans-serif'
						},
						formatter: function(){
							if(typeof(this.value) === "number" && this.value % 1 !== 0){
								return this.value.toFixed(2);
							}
							return this.value;
						}
					}
				}
			},config);

			if(config.xAxis.categories){
				var categories = spreadsheet.getCellRangeValues(config.xAxis.categories);
				(function(categories){
					for(var x = 0; x < categories.length; x++){
						categories[x] = categories[x];
					}
				})(categories);
				config.xAxis.categories = categories;

			}

			for(var i = 0; i < config.series.length; i++){
				var series = config.series.pop();
				var seriesData = spreadsheet.getCellRange(series.data);
				//highcharts requires numeric data, spreadsheet does not automatically convert
				for(var j = 0; j < seriesData.length; j++){
					var cell = seriesData[j];
					seriesData[j] = {
						id:cell.position,
						y:cell.valueOf()
					};
				}

				config.series.unshift({
					data:seriesData,
					name:spreadsheet.getCellValue(series.name)
				});
			}
			console.log(JSON.stringify(config,function(key,val){
				return key == "id" ? val.toString() : val;
			},2));
			new Highcharts.Chart(config,getOnLoadCallback(spreadsheet));
			//Chart is now appended, so lets clear the interval
			clearInterval(appendChartInterval);
		}
	}

	for(var key in window.spreadsheet){
		if(window.spreadsheet[key]){
			var spreadsheetData = window.spreadsheet[key];
			var selector = '#'+key;
			if(spreadsheetData['adapter'] === 'phpexcel'){
				//PHPEXCEL ADAPTER (Server side prosessing of excel documents)
				$( selector).find('.progressbar').progressbar({
					value: false
				});
				var file = window.spreadsheet[key].file;
				var sheet = window.spreadsheet[key].sheet;
				$.get(
					mw.util.wikiScript(), {
						action: 'ajax',
						rs: 'SpreadsheetAjax::getData',
						rsargs: [file,sheet]
					}
					).done(function(resp){
						json = JSON.parse(resp);
						var cellData = json.data;
						var chartData = json.charts;

					//hack to get slickgrid to render enough cells so charts can be anchored
					(function(chartData,cellData){
						if(chartData){
							for(var x = 0; x < chartData.length; x++){
								var chart = chartData[x];
								if(!cellData[chart.end.cell]){
									cellData[chart.end.cell] = '';
								}
							}
						}
					})(chartData,cellData);

					//Pass data to spreadsheet
					var spreadsheetSlick = createSpreadsheet(selector,cellData);

					if(chartData){
						for(var x = 0; x < chartData.length; x++){
							var chart = chartData[x];
							var hc = convertChartDataToHighCharts(chart,cellData);
							renderChart(hc,spreadsheetSlick);
						}
					}
				});


				}else if(spreadsheetData['adapter'] === 'zip'){
				//ZIP ADAPTER (client side prosessing of excel documents)
				spreadsheetZipAdapter = new SpreadsheetZipAdapter(
					spreadsheetData.url,
					spreadsheetData.sheetIndex,
					function(data){
						console.log(data);
						createSpreadsheet(selector,data);
					},
					function(c,max){
						var progressbar = $(selector).children('.progressbar');
						progressbar.progressbar( "option", {
							value:c/max*100
						});
					}
					);
			}

		}
	}
});