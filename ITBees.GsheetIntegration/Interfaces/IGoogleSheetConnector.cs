namespace ITBees.GsheetIntegration.Interfaces
{
    public interface IGoogleSheetConnector
    {
        GSheet<T> Get<T>(string gsheetId, string sheetName, string range, bool firstRowHeader) where T : IGuidItem;
        GSheet Get(string gsheetId, string sheetName, string range, bool firstRowHeader);

        T Insert<T> (string gsheetId, T item, string range);
        T InsertUnique<T> (string gsheetId, T item, string range, Func<T, bool> uniqueQuery, string uniqueuniqueProperty) where T: IGuidItem;
    }
}