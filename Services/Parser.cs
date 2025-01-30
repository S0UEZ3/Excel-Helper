using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EasyBudget.Services
{
    public class Parser
    {
        public List<string[]> ReadExcelFile(string filePath)
        {
            var rows = new List<string[]>();

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false)) // using guarantee that file will be closed
            {
                var workbookPart = doc.WorkbookPart; // whole book structure (lists, styles, links etc.)
                foreach (var sheet in workbookPart.Workbook.Sheets.Elements<Sheet>())
                {
                    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id); // contains cells data
                    var rowsData = worksheetPart.Worksheet.Descendants<Row>(); // finds all rows on the sheet
                    
                    foreach (var row in rowsData) 
                    {
                        var rowValues = new List<string>(); // store cells data from current row
                        foreach (var cell in row.Elements<Cell>()) // get all cells in row
                        {
                            rowValues.Add(GetCellValue(workbookPart, cell)); 
                        }
                        rows.Add(rowValues.ToArray());
                    }
                }
                
                /*var sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>();*/ // take first sheet from collection of all lists
            }
            return rows;
        }

        private static string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            if (cell.CellValue == null)
            {
                return string.Empty;
            }

            string value = cell.CellValue.Text;
            
            if (cell.DataType?.Value == CellValues.SharedString)
            {
                try
                {
                    if (workbookPart.SharedStringTablePart?.SharedStringTable == null)
                    {
                        return value; 
                    }
                    
                    int index = int.Parse(value);
                    return workbookPart.SharedStringTablePart.SharedStringTable.ChildElements[index].InnerText;
                }
                catch (FormatException)
                {
                    Console.WriteLine($"Warning: Unable to parse shared string index '{value}'.");
                }
                catch (ArgumentOutOfRangeException)
                {
                    Console.WriteLine($"Warning: Shared string index '{value}' is out of range.");
                }
            }

            return value;
        }
    }
}