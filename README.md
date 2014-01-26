ExcelImporter
=============

var importer = MakeExcelImporter();

importer.Import(@"d:\excelfile.xlsx")

ExcelImporter MakeExcelImporter()
{
	return new ExcelImporter(new XlsWorkBookReader(), new XlsxWorkBookReader());
}

