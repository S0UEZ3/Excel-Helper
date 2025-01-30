using RowModel = EasyBudget.Models.Row;

namespace EasyBudget.Services
{
    public class SearchSystem(Parser parser)
    {
        private readonly Parser _parser = parser ?? throw new ArgumentNullException(nameof(parser));

        public List<RowModel> FindRow(string directoryPath, string searchTerm)
        {
            var matchingRows = new List<RowModel>();
            if (matchingRows == null)
            {
                throw new ArgumentNullException(nameof(matchingRows));
            }

            foreach (var file in Directory.GetFiles(directoryPath, "*.xlsx"))
            {
                var rows = _parser.ReadExcelFile(file); 
                foreach (var row in rows)
                {
                    if (row.Any(cell => cell.Equals(searchTerm, StringComparison.OrdinalIgnoreCase)))
                    {
                        matchingRows.Add(new RowModel(row, file));
                    }
                }
            }
            return matchingRows;
        }
    }
}