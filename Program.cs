// See https://aka.ms/new-console-template for more information
using ExcelExport.Export;
using ExcelExport.ViewData;

List<ExportViewData> rowItems = new List<ExportViewData>()
{
    new ExportViewData{Id=1, Name="Ball"},
    new ExportViewData{ Id=2, Name="Hat"}
};
Console.WriteLine("Hello, World!");
ExcelExportLib excel = new ExcelExportLib();
var responseStream = excel.CreateExcelDocument<ExportViewData>(rowItems, "2A41E4");
responseStream.Seek(0, SeekOrigin.Begin);
//write to file
FileStream file = new FileStream("c:\\Venkat\\file.xlsx", FileMode.Create, FileAccess.Write);
responseStream.WriteTo(file);
file.Close();
responseStream.Close();