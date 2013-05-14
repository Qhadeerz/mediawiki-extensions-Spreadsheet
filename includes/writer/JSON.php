<?php
/**
 * @author Kim Eik
 */

class PHPExcel_Writer_JSON {

	private $phpExcel;


	public function __construct ( PHPExcel $phpExcel ) {
		$this->phpExcel = $phpExcel;
	}

	public function toJson () {
		// Fetch sheet
		$sheet = $this->phpExcel->getSheet( 0 );

		//$columnDimensions = array();
		//foreach ( $sheet->getColumnDimensions() as $columnDimension ) {
		//	$columnDimensions[ ] = array(
		//		'index' => $columnDimension->getColumnIndex(),
		//		'width' => ( $columnDimension->getWidth() * 24 ) //http://office.microsoft.com/en-001/excel-help/measurement-units-and-rulers-in-excel-HP001151724.aspx
		//	);
		//}

		//initialize output json
		$json = array(
			//'cols'             => PHPExcel_Cell::columnIndexFromString( $sheet->getHighestDataColumn() ),
			//'rows'             => $sheet->getHighestDataRow(),
			'data' => array(),
			//'columnDimensions' => $columnDimensions,
			);

		//create merged cells map
		$mergedCells = array_keys( $sheet->getMergeCells() );
		$cellColspan = array();
		foreach ( $mergedCells as $mc ) {
			$range = PHPExcel_Cell::getRangeBoundaries( $mc );
			$first = PHPExcel_Cell::columnIndexFromString( $range[ 0 ][ 0 ] );
			$last = PHPExcel_Cell::columnIndexFromString( $range[ 1 ][ 0 ] );
			$cellColspan[ $range[ 0 ][ 0 ] . $range[ 0 ][ 1 ] ] = $last - $first + 1;
		}
		unset( $mergedCells );


		//TODO make max limit configurable
		$toRow = min( 2000, $sheet->getHighestRow() );
		$toCol = min( 100, PHPExcel_Cell::columnIndexFromString( $sheet->getHighestColumn() ) );
		for ( $row = 1; $row <= $toRow; $row++ ) {
			for ( $col = 0; $col < $toCol; $col++ ) {
				$cell = $sheet->getCellByColumnAndRow( $col, $row );
				$formula = $cell->getValue();

				if ( $formula === null || strlen( $formula ) === 0 ) {
					continue;
				}

				//remove xlfn prefix for 2010 excel formulas
				$needle = '_xlfn.';
				if ( $formula[ 0 ] === '=' && strpos( $formula, $needle ) !== false ) {
					$formula = str_replace( $needle, '', $formula );
				}

				$json[ 'data' ][ $cell->getCoordinate() ][ 'formula' ] = $formula;

				if ( $cell->hasDataValidation() ) {
					$json[ 'data' ][ $cell->getCoordinate() ][ 'dataValidation' ] = $this->getDataValidation( $cell->getDataValidation() );
				}

				if ( array_key_exists( $cell->getCoordinate(), $cellColspan ) ) {
					$json[ 'data' ][ $cell->getCoordinate() ][ "metadata" ][ 'colspan' ] = $cellColspan[ $cell->getCoordinate() ];
				}


			}
		}

		//get charts
		$charts = $sheet->getChartCollection();
		foreach ( $charts as $chart ) {
			$json[ 'charts' ][ ] = array(
				'title' => is_null( $chart->getTitle() ) ? '' : $this->getChartTitle( $chart->getTitle() ),
				'data'  => $this->getPlotArea( $chart->getPlotArea() ),
				'start' => $chart->getTopLeftPosition(),
				'end'   => $chart->getBottomRightPosition(),
				/*'labels' => array(
					'xAxis' =>is_null($chart->getXAxisLabel()) ? '' : $chart->getXAxisLabel()->getCaption()->getPlainText(),
					'yAxis' => is_null($chart->getYAxisLabel()) ?  '' : $chart->getYAxisLabel()->getCaption()->getPlainText(),
					),*/
			);
		}

		return json_encode( $json );
	}

	private function getPlotArea ( PHPExcel_Chart_PlotArea $plotArea ) {
		return array(
			//'layout'    => $this->getLayout( $plotArea->getLayout() ),
			'plotGroup' => $this->getPlotGroupCollection( $plotArea->getPlotGroup() ),
			);
	}

	private function getLayout ( PHPExcel_Chart_Layout $layout ) {
		return array(
			'target'          => $layout->getLayoutTarget(),
			'xmode'           => $layout->getXMode(),
			'ymode'           => $layout->getYMode(),
			'xpos'            => $layout->getXPosition(),
			'ypos'            => $layout->getYPosition(),
			'width'           => $layout->getWidth(),
			'height'          => $layout->getHeight(),
			'showLegendKey'   => $layout->getShowLegendKey(),
			'showValue'       => $layout->getShowVal(),
			'showCatName'     => $layout->getShowCatName(),
			'showSerName'     => $layout->getShowSerName(),
			'showPercent'     => $layout->getShowPercent(),
			'showBubbleSize'  => $layout->getShowBubbleSize(),
			'showLeaderLines' => $layout->getShowLeaderLines(),
			);
	}

	private function getPlotGroupCollection ( array $plotGroups ) {
		$result = array();
		foreach ( $plotGroups as $plotGroup ) {
			$result[ ] = $this->getPlotGroup( $plotGroup );
		}

		return $result;
	}

	private function getPlotGroup ( PHPExcel_Chart_DataSeries $plotGroup ) {
		return array(
			'type'       => $plotGroup->getPlotType(),
			'grouping'   => $plotGroup->getPlotGrouping(),
			'direction'  => $plotGroup->getPlotDirection(),
			'style'      => $plotGroup->getPlotStyle(),
			'smoothLine' => $plotGroup->getSmoothLine(),
			'data'       => $this->getPlotData( $plotGroup ),
			'seriesCount' => $plotGroup->getPlotSeriesCount(),
			);

	}

	private function getDataSeriesValuesCollection ( array $dataSeriesValues ) {
		$result = array();
		foreach ( $dataSeriesValues as $key => $dataSeriesValue ) {
			$result[ $key ] = $this->getDataSeriesValue( $dataSeriesValue );
		}

		return $result;
	}

	private function getDataSeriesValue ( PHPExcel_Chart_DataSeriesValues $dataSeriesValue ) {
		return array(
			'dataType'   => $dataSeriesValue->getDataType(),
			'dataSource' => $dataSeriesValue->getDataSource(),
			'format'     => $dataSeriesValue->getFormatCode(),
			'marker'     => $dataSeriesValue->getPointMarker(),
			);
	}

	public function getChartTitle ( PHPExcel_Chart_Title $title ) {
		$captions = $title->getCaption();
		return $this->getRichText( $captions[ 0 ] );
	}

	private function getRichText ( PHPExcel_RichText $text ) {
		$result = array();
		foreach ( $text->getRichTextElements() as $textElement ) {
			$result[ ] = $this->getRichTextElement( $textElement );
		}

		return $result;
	}

	private function getRichTextElement ( PHPExcel_RichText_TextElement $text ) {
		return array(
			'text' => $text->getText(),
			'font' => $this->getFont( $text->getFont() ),
			);
	}

	private function getFont ( $font ) {
		return array(
			'underline'     => $font->getUnderline(),
			'bold'          => $font->getBold(),
			'name'          => $font->getName(),
			'size'          => $font->getSize(),
			'color'         => $font->getColor()->getRGB(),
			'italic'        => $font->getItalic(),
			'strikethrough' => $font->getStrikethrough(),
			'subscript'     => $font->getSubScript(),
			'superscript'   => $font->getSuperScript(),
			);
	}

	private function getPlotData ( PHPExcel_Chart_DataSeries $plotGroup ) {
		$ordering = $plotGroup->getPlotOrder();
		$data = array(
			'label'    => $this->getDataSeriesValuesCollection( $plotGroup->getPlotLabels() ),
			'category' => $this->getDataSeriesValuesCollection( $plotGroup->getPlotCategories() ),
			'value'    => $this->getDataSeriesValuesCollection( $plotGroup->getPlotValues() ),
			);
		$result = array();
		foreach ( $ordering as $oKey => $oVal ) {
			foreach ( $data as $dKey => $dVal ) {
				if ( array_key_exists( $oKey, $dVal ) ) {
					if ( !isset( $result[ $oVal ] ) ) {
						$result[ $oVal ] = array();
					}
					$result[ $oVal ][ $dKey ] = $dVal[ $oKey ];
				}
			}
		}

		return $result;
	}

	private function getDataValidation ( PHPExcel_Cell_DataValidation $validation ) {
		return array(
			'type'     => $validation->getType(),
			'operator'         => $validation->getOperator() === '' ? 'between' : $validation->getOperator(),
			'args' => array(
				$validation->getFormula1(),
				$validation->getFormula2(),
				),
			'options' => array(
				'allowBlank'       => $validation->getAllowBlank(),
				'errorText'        => $validation->getError(),
				'errorType'        => $validation->getErrorStyle(),
				'errorTitle'       => $validation->getErrorTitle(),
				'promptText'       => $validation->getPrompt(),
				'promptTitle'      => $validation->getPromptTitle(),
				'showDropDown'     => $validation->getShowDropDown(),
				'showErrorMessage' => $validation->getShowErrorMessage(),
				'showInputMessage' => $validation->getShowInputMessage(),				
				),
			);
	}

}