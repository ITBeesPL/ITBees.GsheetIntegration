using Google.Apis.Sheets.v4.Data;
using ITBees.Models.Languages;

namespace ITBees.GsheetIntegration.Interfaces
{
    public interface IGoogleSheetConnector
    {
        GSheet Get(string gsheetId, string sheetName, string range, bool firstRowHeader);

        T Insert<T>(string gsheetId, T item, string range);
        T InsertUnique<T>(string gsheetId, T item, string range, Func<T, bool> uniqueQuery, string uniqueuniqueProperty, Language lang) where T : class, IGuidItem;

        Spreadsheet CreateSpreadsheet(string newGoogleSpreadsheetOwnerEmail, string worksheetName);
        Spreadsheet CreateSpreadsheet(string newGoogleSpreadsheetOwnerEmail, string worksheetName, string copyTemplateSheetId);
    }
}