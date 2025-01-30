namespace EasyBudget.Helpers
{
    public class UserInteractions
    {
        public (string directoryPath, string searchTerm, bool createNewFile, string outputFilePath) GetUserInput(string[] args)
        {
            if (args.Length < 4)
            {
                ErrorManager.HandleInvalidInputError();
            }

            var directoryPath = args[0];
            var searchTerm = args[1];
            var writeMode = args[2];
            var outputFilePath = args[3];
            
            if (!Directory.Exists(directoryPath))
            {
                ErrorManager.HandleFileNotFoundError(directoryPath);
            }

            bool createNewFile = writeMode.ToLower() == "new";

            return (directoryPath, searchTerm, createNewFile, outputFilePath);
        }

        public static void DisplayResult(string message)
        {
            Console.WriteLine(message);
        }
    }
}

// dotnet run -- "C:\Users\Admin\Documents\TestData" "item1" "new" "C:\Users\Admin\Documents\TestDataResults\output.xlsx"