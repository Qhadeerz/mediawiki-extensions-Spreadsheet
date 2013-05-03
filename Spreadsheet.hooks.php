<?php

class SpreadsheetHooks {

	public static function init(Parser &$parser) {
		$parser->setHook('spreadsheet', 'SpreadsheetHooks::renderFromTag');
		return true;
	}

	public static function renderFromTag($input, array $args, Parser $parser, PPFrame $frame) {
		$file = wfLocalFile($args['file']);
		$style = isset($args['style']) ? $args['style'] : 'height: 600px;'; //css style parameter definition
		$class = isset($args['class']) ? $args['class'] : 'spreadsheet-container'; //space separated list of css classes
		if($file->exists()){
			$parser->getOutput()->addModules('spreadsheet.core');


			// $output = json_encode( array(
			// 	'adapter' => 'zip',
			// 	'url' => $file->getUrl(),
			// 	'sheetIndex' => $sheetIndex,
			// 	));
			// $parser->getOutput()->addModules('spreadsheet.core');
			// $output = self::getOutputHTML($parser->getRandomString(),$output);

			return array(self::getOutputHTML($parser->getRandomString(),array(
				'adapter' => 'phpexcel',
				'file' => $args['file'],
				'sheet' => isset($args['sheet']) ? $args['sheet'] : '0',
			),$class,$style), 'noparse' => true, 'isHTML' => true);
		}else{
			return "file doesn't exist";
		}

	}

	public static function registerUnitTests(&$files) {
		$testDir = dirname(__FILE__) . '/test/phpunit/';
		$testFiles = scandir($testDir);
		foreach ($testFiles as $testFile) {
			$absoluteFile = $testDir . $testFile;
			if (is_file($absoluteFile)) {
				$files[] = $absoluteFile;
			}
		}
		return true;
	}

	public static function registerQunitTests(array &$testModules, \ResourceLoader &$resourceLoader ){
		$testModules['qunit']['spreadsheet.tests'] = array(
			'scripts' => array( 'test/qunit/spreadsheet.test.js' ),
			'dependencies' => array( 'spreadsheet.core' ),
			'localBasePath' => dirname( __FILE__ ),
			'remoteExtPath' => 'Spreadsheet',
		);
		return true;
	}

	private static function getOutputHTML($id,$data,$class,$style){
		$data = json_encode($data);
		return <<<HTML
			<script type="text/javascript">
				<!--
				if(window.spreadsheet === undefined){
					window.spreadsheet = {};
				}
				window.spreadsheet['$id'] = $data;
				-->
			</script>
			<div id="$id" class="$class" style="$style">
				<div class="progressbar"></div>
			</div>
HTML;
	}
}
