using System.Reflection;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using ITBees.GsheetIntegration.Interfaces;
using ITBees.GsheetIntegration.Tools;
using ITBees.Models.Languages;

namespace ITBees.GsheetIntegration
{
    public class GoogleSheetConnector : IGoogleSheetConnector
    {
        private readonly GoogleSheetsHelper _googleSheetsHelper;
        private readonly GoogleSheetConnector _googleSheetConnector;
        private dynamic _gsheet;

        public GoogleSheetConnector(GoogleSheetsHelper googleSheetsHelper)
        {
            _googleSheetsHelper = googleSheetsHelper;
        }
        public GSheet<T> Get<T>(string sheetId, string sheetName, string range, bool firstRowHeader, Language lang) where T : class, IGuidItem
        {
            var values = _googleSheetsHelper.Service.Spreadsheets.Values;
            var request = values.Get(sheetId, range);
            var response = request.Execute();

            var result = GSheetTypeConverter.Convert<T>(response, firstRowHeader, lang);

            return result;
        }

        public GSheet Get(string sheetId, string sheetName, string range, bool firstRowHeader)
        {
            var values = _googleSheetsHelper.Service.Spreadsheets.Values;
            var request = values.Get(sheetId, range);
            var response = request.Execute();

            var result = GSheetTypeConverter.Convert(response, firstRowHeader);

            return result;
        }

        public T Insert<T>(string sheetId, T item, string range)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// If specified unique query will find item already saved in google sheet, it will return its GUID, and nothing else will be updated.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="item"></param>
        /// <param name="range">Should be whole sheet range like : CompanyPartner!A:F</param>
        /// <returns></returns>
        public T InsertUnique<T>(string gsheetId, T item, string range, Func<T, bool> uniqueQuery, string uniqueProperty, Language lang) where T : class, IGuidItem
        {
            GSheet<T> gsheet;
            if (_gsheet == null)
            {
                _gsheet = Get<T>(gsheetId, "", range, true, lang);
                gsheet = _gsheet;
            }
            else
            {
                if (_gsheet.GetType() == typeof(GSheet<T>))
                {
                    gsheet = _gsheet;
                }
                else
                {
                    gsheet = null;
                    _gsheet = Get<T>(gsheetId, "", range, true, lang);
                    gsheet = _gsheet;
                }

            }

            IList<T> currentData = gsheet.Data;
            var alreadySavedItem = currentData.FirstOrDefault(uniqueQuery);
            if (alreadySavedItem != null)
            {
                item.Guid = alreadySavedItem.Guid;
                return item;
            }

            if (item.Guid == Guid.Empty)
            {
                item.Guid = Guid.NewGuid();
            }

            var valueRange = new ValueRange();

            var objectList = new List<object>();
            var i = 0;
            foreach (var column in gsheet.ColumnsIndexes.OrderBy(x => x.Key))
            {
                PropertyInfo pi = null;
                if (i == 0)
                {
                    pi = item.GetType().GetProperties().FirstOrDefault(x => x.Name == "Guid");
                }
                else
                {
                    pi = item.GetType().GetProperties().FirstOrDefault(x => x.Name == column.Value.Replace("-","") + $"_{i}");
                }
                var value = pi.GetValue(item);
                value = value == null ? string.Empty : value.ToString();
                if (pi.PropertyType == typeof(string))
                {
                    objectList.Add(value);
                }
                else if (pi.PropertyType == typeof(bool))
                {
                    objectList.Add(bool.Parse(value.ToString()));
                }
                else if (pi.PropertyType == typeof(Guid))
                {
                    objectList.Add((value.ToString()));
                }
                else
                {
                    objectList.Add(string.Empty);
                }

                i++;
            }


            valueRange.Values = new List<IList<object>>() { objectList };

            var appendRequest =
                _googleSheetsHelper.Service.Spreadsheets.Values.Append(valueRange, gsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

            var appendResponse = ExecuteRequest<T>(appendRequest);

            return item;
        }

        public Spreadsheet CreateSpreadsheet(string newGoogleSpreadsheetOwnerEmail, string worksheetName)
        {
            return _googleSheetsHelper.CreateSpreadsheet(newGoogleSpreadsheetOwnerEmail, worksheetName, String.Empty);
        }

        public Spreadsheet CreateSpreadsheet(string newGoogleSpreadsheetOwnerEmail, string worksheetName, string copyTemplateSheetId)
        {
            return _googleSheetsHelper.CreateSpreadsheet(newGoogleSpreadsheetOwnerEmail, worksheetName, copyTemplateSheetId);
        }

        private static AppendValuesResponse ExecuteRequest<T>(SpreadsheetsResource.ValuesResource.AppendRequest appendRequest) where T : IGuidItem
        {
            try
            {
                return appendRequest.Execute();
            }
            catch (Exception e)
            {
                if (e.Message.Contains("Limit"))
                {
                    System.Threading.Thread.Sleep(80 * 1000);
                    return ExecuteRequest<T>(appendRequest);
                }
            }

            return null;
        }
    }
}