<?php

/**
 * GridFieldExportToExcelButton
 *
 * Adds an "Export list" button to the bottom of a GridField.
 *
 * Most code is the same as GridFieldExportButton
 *
 * @author     Martijn van Nieuwenhoven <info@axyrmedia.nl>
 */
class GridFieldExportToExcelButton extends GridFieldExportButton {

	/**
	 * Place the export button in a <p> tag below the field
	 */
	public function getHTMLFragments($gridField) {
		$button = new GridField_FormAction(
			$gridField, 
			'export', 
			_t('TableListField.EXCELEXPORT', 'Export to Excel'),
			'export', 
			null
		);
		$button->setAttribute('data-icon', 'download-csv');
		$button->addExtraClass('no-ajax');
		return array(
			$this->targetFragment => '<p class="grid-csv-button">' . $button->Field() . '</p>',
		);
	}

	/**
	 * Handle the export, for both the action button and the URL
 	 */
	public function handleExport($gridField, $request = null) {
		$now = date("d-m-Y-H-i");
		$title = str_replace(' ', '-', strtolower(singleton($gridField->getModelClass())->singular_name()));
		$fileName = "export-$title-$now.xlsx";

		if($fileData = $this->generateExportFileData($gridField)){
			return SS_HTTPRequest::send_file($fileData, $fileName, 'application/xslt+xml');
		}
	}

	/**
	 * Generate export fields for EXCEL.
	 *
	 * @param GridField $gridField
	 * @return array
	 */
	public function generateExportFileData($gridField) {
		
		$excelColumns = ($this->exportColumns)
			? $this->exportColumns
			: singleton($gridField->getModelClass())->summaryFields();
		
		$objPHPExcel = new PHPExcel();
		$worksheet = $objPHPExcel->getActiveSheet()->setTitle(singleton($gridField->getModelClass())->i18n_singular_name());
		
		$col = 'A';
		foreach($excelColumns as $columnSource => $columnHeader) {
			$heading = (!is_string($columnHeader) && is_callable($columnHeader)) ? $columnSource : $columnHeader;
			$worksheet->setCellValue($col.'1', $heading);
			$col++;
		}
			
		$worksheet->freezePane('A2');
		
		$items = $gridField->getManipulatedList();

		// @todo should GridFieldComponents change behaviour based on whether others are available in the config?
		foreach($gridField->getConfig()->getComponents() as $component){
			if($component instanceof GridFieldFilterHeader || $component instanceof GridFieldSortableHeader) {
				$items = $component->getManipulatedData($gridField, $items);
			}
		}

		$row = 2;
		foreach($items->limit(null) as $item) {
			if(!$item->hasMethod('canView') || $item->canView()) {
				$columnData = array();
				$col = 'A';
				foreach($excelColumns as $columnSource => $columnHeader) {
					if(!is_string($columnHeader) && is_callable($columnHeader)) {
						if($item->hasMethod($columnSource)) {
							$relObj = $item->{$columnSource}();
						} else {
							$relObj = $item->relObject($columnSource);
						}

						$value = $columnHeader($relObj);
						$worksheet->getCell($col.$row)->setValueExplicit($value, PHPExcel_Cell_DataType::TYPE_STRING);
					} else {
						$component = $item;
						$value = $gridField->getDataFieldValue($item, $columnSource);

						if(strpos($columnSource, '.') !== false) {
							$relations = explode('.', $columnSource);
							foreach($relations as $relation) {
								if($component->hasMethod($relation)) {
									$component = $component->$relation();
								} 
								elseif($component instanceof SS_List) {
									$component = $component->relation($relation);
								} 
								elseif($component instanceof DataObject && ($dbObject = $component->obj($relation))) {
									$component = $dbObject;
								}
							}
						}
						elseif($component instanceof DataObject && ($dbObject = $component->obj($columnSource))) {
							$component = $dbObject;
						}

						if(!$value) {
							$component = $item;
							$value = $gridField->getDataFieldValue($item, $columnHeader);

							if(strpos($columnHeader, '.') !== false) {
								$relations = explode('.', $columnHeader);
								foreach($relations as $relation) {
									if($component->hasMethod($relation)) {
										$component = $component->$relation();
									} 
									elseif($component instanceof SS_List) {
										$component = $component->relation($relation);
									} 
									elseif($component instanceof DataObject && ($dbObject = $component->obj($relation))) {
										$component = $dbObject;
									}
								}
							}
							elseif($component instanceof DataObject && ($dbObject = $component->obj($columnHeader))) {
								$component = $dbObject;
							}
						}
						
						if($component && ($component instanceof Decimal || $component instanceof Float || $component instanceof Int)){
							$worksheet->getCell($col.$row)->setValue($value);
						}
						else{
							$worksheet->getCell($col.$row)->setValueExplicit($value, PHPExcel_Cell_DataType::TYPE_STRING);
						}
					}
					
					$col++;
					
				}
				
				$row++;
			}
			if($item->hasMethod('destroy')) {
				$item->destroy();
			}
		}
		
		$writer = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		ob_start();
		$writer->save('php://output');
		$data = ob_get_clean();
		return $data;
	}
}
