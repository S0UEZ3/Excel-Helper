using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

// Alias to resolve ambiguity between custom Row and OpenXML Row
using RowModel = EasyBudget.Models.Row;

namespace EasyBudget.Services
{
    public class Writer
    {
        public void WriteRows(List<RowModel> rows, string outputFilePath, bool createNewFile)
    {
        using (SpreadsheetDocument doc = createNewFile
            ? SpreadsheetDocument.Create(outputFilePath, SpreadsheetDocumentType.Workbook)
            : SpreadsheetDocument.Open(outputFilePath, true))
        {
            var workbookPart = doc.WorkbookPart ?? doc.AddWorkbookPart();
            
            // Создаем Workbook, если его нет
            if (workbookPart.Workbook == null)
                workbookPart.Workbook = new Workbook();

            var sheetPart = workbookPart.WorksheetParts.FirstOrDefault() ?? workbookPart.AddNewPart<WorksheetPart>();

            // Создаем Worksheet, если его нет
            if (sheetPart.Worksheet == null)
                sheetPart.Worksheet = new Worksheet(new SheetData());

            var sheetData = sheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                sheetData = new SheetData();
                sheetPart.Worksheet.AppendChild(sheetData);
            }

            // Создаем лист в Workbook, если он не добавлен
            if (workbookPart.Workbook.Sheets == null)
                workbookPart.Workbook.AppendChild(new Sheets());

            if (!workbookPart.Workbook.Sheets.Elements<Sheet>().Any())
            {
                var sheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(sheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };
                workbookPart.Workbook.Sheets.Append(sheet);
            }

            // Добавляем данные в sheetData
            foreach (var rowModel in rows)
            {
                var newRow = new Row();
                foreach (var cellValue in rowModel.RowData)
                {
                    newRow.AppendChild(new Cell { CellValue = new CellValue(cellValue), DataType = CellValues.String });
                }
                sheetData.AppendChild(newRow);
            }

            sheetPart.Worksheet.Save();
            workbookPart.Workbook.Save();
        }
    }
        
        /*private void CreateNewExcelFile(string outputFilePath, string[] rowData)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(outputFilePath, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                var sheets = doc.WorkbookPart?.Workbook.AppendChild(new Sheets());
                sheets?.Append(new Sheet
                {
                    Id = doc.WorkbookPart?.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Output"
                });

                AddRowToSheet(worksheetPart, rowData);
                workbookPart.Workbook.Save();
            }
        }*/

        /*private void AppendToExistingExcelFile(string outputFilePath, string[] rowData)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(outputFilePath, true))
            {
                var worksheetPart = doc.WorkbookPart?.WorksheetParts.First();
                if (worksheetPart != null)
                {
                    AddRowToSheet(worksheetPart, rowData);
                }
                doc.WorkbookPart?.Workbook.Save();
            }
        }*/

        /*private static void AddRowToSheet(WorksheetPart worksheetPart, string[] rowData)
        {
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var newRow = new Row(); // Explicitly specify OpenXML Row.

            foreach (var cellValue in rowData)
            {
                newRow.Append(new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(cellValue)
                });
            }

            sheetData?.Append(newRow);
        }*/
    }
}
