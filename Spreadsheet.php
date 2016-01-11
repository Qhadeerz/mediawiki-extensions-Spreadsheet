<?php

if (!defined('MEDIAWIKI')) {
	die('Not an entry point.');
}

$wgExtensionFunctions[] = function () {
	if (!class_exists( "PHPExcel" )) {
		die('Spreadsheet extensions requires PHPExcel extension');
	}
};

$wgExtensionCredits['parserhook'][] = array(
	'path' => __FILE__,
	'name' => 'Spreadsheet',
	'version' => '0.2.0',
	'author' => 'Kim Eik',
	'url' => 'https://www.mediawiki.org/wiki/Extension:Spreadsheet',
	'descriptionmsg' => 'spreadsheet-desc'
);


// Specify the function that will initialize the parser function.
$wgHooks['ParserFirstCallInit'][] = 'SpreadsheetHooks::init';

// Sepcify phpunit tests
//$wgHooks['UnitTestsList'][] = 'SpreadsheetHooks::registerUnitTests';

//Register qunit tests
//$wgHooks['ResourceLoaderTestModules'][] = 'SpreadsheetHooks::registerQunitTests';

//Autoload hooks
$wgAutoloadClasses['SpreadsheetHooks'] = dirname(__FILE__) . '/Spreadsheet.hooks.php';

//Autoload ajax
$wgAutoloadClasses['SpreadsheetAjax'] = dirname(__FILE__) . '/Spreadsheet.ajax.php';

//Autoload classes
$wgSpreadsheetIncludes = dirname(__FILE__) . '/includes';
$wgAutoloadClasses['PHPExcel_Writer_JSON'] = $wgSpreadsheetIncludes . '/writer/JSON.php';

//i18n
$wgMessagesDirs['Spreadsheet'] = __DIR__ . '/i18n';
$wgExtensionMessagesFiles['Spreadsheet'] = dirname(__FILE__) . '/Spreadsheet.i18n.php';

//Resources
$wgResourceModules['spreadsheet.core'] = array(
	'scripts' => array(
		'lib/spreadsheet-js/lib/efp/efp.js',
		'lib/spreadsheet-js/lib/slickgrid/lib/jquery.event.drag-2.0.min.js',
		'lib/spreadsheet-js/lib/slickgrid/plugins/slick.autotooltips.js',
		'lib/spreadsheet-js/lib/slickgrid/plugins/slick.cellrangedecorator.js',
		'lib/spreadsheet-js/lib/slickgrid/plugins/slick.cellrangeselector.js',
		'lib/spreadsheet-js/lib/slickgrid/plugins/slick.cellselectionmodel.js',
		'lib/spreadsheet-js/lib/slickgrid/slick.core.js',
		'lib/spreadsheet-js/lib/slickgrid/slick.grid.js',
		'lib/spreadsheet-js/lib/highcharts/js/highcharts.src.js',
		'lib/spreadsheet-js/lib/highcharts/js/modules/exporting.src.js',
		'lib/spreadsheet-js/spreadsheet.core.js',
		'lib/spreadsheet-js/spreadsheet.slick.js',
		//'lib/zip/WebContent/zip.js',
		//'lib/zip/WebContent/zip-ext.js',
		//'lib/zip/WebContent/deflate.js',
		//'lib/zip/WebContent/inflate.js',
		//'js/spreadsheet.zip.adapter.js',
		'js/spreadsheet.bootstrap.js'
	),
	'styles' => array(
		'css/spreadsheet.css',
		'css/slick.grid.css'
	),
	'dependencies' => array(
		'jquery.ui.core',
		'jquery.ui.dialog',
		'jquery.ui.autocomplete',
		'jquery.ui.progressbar',
		'jquery.effects.core',
	),
	'group' => 'spreadsheet',
	'localBasePath' => dirname(__FILE__),
	'remoteExtPath' => 'Spreadsheet'
);

$wgAjaxExportList[] = "SpreadsheetAjax::getData";
