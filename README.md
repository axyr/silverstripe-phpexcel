silverstripe-phpexcel
=====================

Export DataObjects to an Excel file

:::php
	$gridField = new GridField('MyDataObjects', 'MyDataObjects', MyDataObject::get(), 
			GridFieldConfig_RecordEditor::create()
			->addComponent(new GridFieldExportToExcelButton())
		);
