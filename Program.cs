using EasyBudget.Helpers;
using EasyBudget.Services;

namespace EasyBudget
{
    internal abstract class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                var userInteractions = new UserInteractions();
                var (directoryPath, searchTerm, createNewFile, outputFilePath) = userInteractions.GetUserInput(args);

                var parser = new Parser(); 
                var searchSystem = new SearchSystem(parser); 

                var matchingRows = searchSystem.FindRow(directoryPath, searchTerm);

                if (matchingRows is { Count: 0 })
                {
                    ErrorManager.HandleItemNotFoundError(searchTerm);
                }

                var writer = new Writer();
                writer.WriteRows(matchingRows, outputFilePath, createNewFile);

                UserInteractions.DisplayResult($"Item '{searchTerm}' was successfully added to '{outputFilePath}'");
            }
            catch (Exception ex)
            {
                ErrorManager.HandleGeneralError(ex);
            }
        }
    }
}

// example run command:
// dotnet run -- "C:\Users\Admin\Documents\TestData" "item1" "new" "C:\Users\Admin\Documents\TestDataResults\output.xlsx"