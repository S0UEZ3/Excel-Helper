namespace EasyBudget.Helpers
{
    public static class ErrorManager
    {
        public static void HandleFileNotFoundError(string filePath)
        {
            throw new FileNotFoundException($"Error: File not found at '{filePath}'", filePath);
        }

        public static void HandleItemNotFoundError(string searchTerm)
        {
            throw new KeyNotFoundException($"Error: Item '{searchTerm}' not found in any Excel file.");
        }

        public static void HandleDuplicateItemError(string searchTerm)
        {
            throw new InvalidOperationException($"Error: Item '{searchTerm}' is already present in the output file.");
        }

        public static void HandleInvalidInputError()
        {
            throw new ArgumentException("Error: Invalid input. Please check your command-line arguments.");
        }

        public static void HandleGeneralError(Exception ex)
        {
            throw new Exception($"Error: {ex.Message}", ex);
        }
    }
}