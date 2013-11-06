silverstripe-phpexcel
=====================

Same as GridFieldExportButton, but exports data to an Excel file.

	:::php
	$gridField = new GridField('MyDataObjects', 'MyDataObjects', MyDataObject::get(), 
			GridFieldConfig_RecordEditor::create()
			->addComponent(new GridFieldExportToExcelButton())
		);

## Requirements
* PHPExcel https://github.com/PHPOffice/PHPExcel

This will install when you use composer.
