// Row.cs - Represents a structured data model for rows in Excel.
namespace EasyBudget.Models
{
    public class Row
    {
        public string[] RowData { get; }
        public string SourceFilePath { get; }
        
        public Row(string[] rowData, string sourceFilePath)
        {
            RowData = rowData ?? throw new ArgumentNullException(nameof(rowData));
            SourceFilePath = sourceFilePath ?? throw new ArgumentNullException(nameof(sourceFilePath));
        }

        // Method to add the source file path as a new column to the row data.
        public string[] GetRowWithSourceFile()
        {
            var result = new string[RowData.Length + 1];
            Array.Copy(RowData, result, RowData.Length);
            result[RowData.Length] = SourceFilePath;
            return result;
        }
    }
}