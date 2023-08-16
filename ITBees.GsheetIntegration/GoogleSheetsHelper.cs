using System.Dynamic;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using ITBees.GsheetIntegration.Models;

namespace ITBees.GsheetIntegration;

public class GoogleSheetsHelper
{
    public string ApplicationName { get; }
    public SheetsService Service { get; set; }
    static readonly string[] Scopes = { DriveService.ScopeConstants.Drive, SheetsService.Scope.Spreadsheets };

    public GoogleSheetsHelper(string credentialFilePath, string applicationName)
    {
        ApplicationName = applicationName;
        Credential = GetCredentialsFromFile(credentialFilePath);
        Service = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = Credential,
            ApplicationName = applicationName
        });
    }

    public GoogleCredential Credential { get; set; }

    private GoogleCredential GetCredentialsFromFile(string filePath)
    {
        GoogleCredential credential;
        using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
        }

        return credential;
    }

    private readonly SheetsService _sheetsService;
    private readonly string _spreadsheetId;



    public List<ExpandoObject> GetDataFromSheet(GoogleSheetParameters googleSheetParameters)
    {
        googleSheetParameters = MakeGoogleSheetDataRangeColumnsZeroBased(googleSheetParameters);
        var range =
            $"{googleSheetParameters.SheetName}!{GetColumnName(googleSheetParameters.RangeColumnStart)}{googleSheetParameters.RangeRowStart}:{GetColumnName(googleSheetParameters.RangeColumnEnd)}{googleSheetParameters.RangeRowEnd}";

        SpreadsheetsResource.ValuesResource.GetRequest request =
            _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);

        var numberOfColumns = googleSheetParameters.RangeColumnEnd - googleSheetParameters.RangeColumnStart;
        var columnNames = new List<string>();
        var returnValues = new List<ExpandoObject>();

        if (!googleSheetParameters.FirstRowIsHeaders)
        {
            for (var i = 0; i <= numberOfColumns; i++)
            {
                columnNames.Add($"Column{i}");
            }
        }

        var response = request.Execute();

        int rowCounter = 0;
        IList<IList<Object>> values = response.Values;
        if (values != null && values.Count > 0)
        {
            foreach (var row in values)
            {
                if (googleSheetParameters.FirstRowIsHeaders && rowCounter == 0)
                {
                    for (var i = 0; i <= numberOfColumns; i++)
                    {
                        columnNames.Add(row[i].ToString());
                    }

                    rowCounter++;
                    continue;
                }

                var expando = new ExpandoObject();
                var expandoDict = expando as IDictionary<String, object>;
                var columnCounter = 0;
                foreach (var columnName in columnNames)
                {
                    expandoDict.Add(columnName, row[columnCounter].ToString());
                    columnCounter++;
                }

                returnValues.Add(expando);
                rowCounter++;
            }
        }

        return returnValues;
    }

    public void AddCells(GoogleSheetParameters googleSheetParameters, List<GoogleSheetRow> rows)
    {
        var requests = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };

        var sheetId = GetSheetId(_sheetsService, _spreadsheetId, googleSheetParameters.SheetName);

        GridCoordinate gc = new GridCoordinate
        {
            ColumnIndex = googleSheetParameters.RangeColumnStart - 1,
            RowIndex = googleSheetParameters.RangeRowStart - 1,
            SheetId = sheetId
        };

        var request = new Request { UpdateCells = new UpdateCellsRequest { Start = gc, Fields = "*" } };

        var listRowData = new List<RowData>();

        foreach (var row in rows)
        {
            var rowData = new RowData();
            var listCellData = new List<CellData>();
            foreach (var cell in row.Cells)
            {
                var cellData = new CellData();
                var extendedValue = new ExtendedValue { StringValue = cell.CellValue };

                cellData.UserEnteredValue = extendedValue;
                var cellFormat = new CellFormat { TextFormat = new TextFormat() };

                if (cell.IsBold)
                {
                    cellFormat.TextFormat.Bold = true;
                }

                cellFormat.BackgroundColor = new Color
                {
                    Blue = (float)cell.BackgroundColor.B / 255,
                    Red = (float)cell.BackgroundColor.R / 255,
                    Green = (float)cell.BackgroundColor.G / 255
                };

                cellData.UserEnteredFormat = cellFormat;
                listCellData.Add(cellData);
            }

            rowData.Values = listCellData;
            listRowData.Add(rowData);
        }

        request.UpdateCells.Rows = listRowData;

        // It's a batch request so you can create more than one request and send them all in one batch. Just use reqs.Requests.Add() to add additional requests for the same spreadsheet
        requests.Requests.Add(request);

        _sheetsService.Spreadsheets.BatchUpdate(requests, _spreadsheetId).Execute();
    }

    private string GetColumnName(int index)
    {
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        var value = "";

        if (index >= letters.Length)
            value += letters[index / letters.Length - 1];

        value += letters[index % letters.Length];
        return value;
    }

    private GoogleSheetParameters MakeGoogleSheetDataRangeColumnsZeroBased(GoogleSheetParameters googleSheetParameters)
    {
        googleSheetParameters.RangeColumnStart = googleSheetParameters.RangeColumnStart - 1;
        googleSheetParameters.RangeColumnEnd = googleSheetParameters.RangeColumnEnd - 1;
        return googleSheetParameters;
    }

    private int GetSheetId(SheetsService service, string spreadSheetId, string spreadSheetName)
    {
        var spreadsheet = service.Spreadsheets.Get(spreadSheetId).Execute();
        var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == spreadSheetName);
        int sheetId = (int)sheet.Properties.SheetId;
        return sheetId;
    }

    public Spreadsheet CreateSpreadsheet(string newGoogleSpreadsheetOwnerEmail, string worksheetName, string templateWorksheetId)
    {
        {
            Spreadsheet spreadsheet = null;

            spreadsheet = new Spreadsheet
            {
                Properties = new SpreadsheetProperties
                {
                    Title = worksheetName
                }
            };

            var createdSpreadsheet = Service.Spreadsheets.Create(spreadsheet).Execute();

            var driveService = new DriveService(new BaseClientService.Initializer
            {
                HttpClientInitializer = Credential,
                ApplicationName = ApplicationName,
            });

            var permission = new Google.Apis.Drive.v3.Data.Permission
            {
                Type = "user",
                Role = "writer",
                EmailAddress = newGoogleSpreadsheetOwnerEmail
            };

            var request = driveService.Permissions.Create(permission, createdSpreadsheet.SpreadsheetId);
            request.Execute();
            var file = driveService.Files.Get(createdSpreadsheet.SpreadsheetId).Execute();
            if (string.IsNullOrEmpty(templateWorksheetId))
                return createdSpreadsheet;

            SpreadsheetsResource.GetRequest getRequest = Service.Spreadsheets.Get(templateWorksheetId);
            getRequest.IncludeGridData = true;
            Spreadsheet templateSpreadSheet = getRequest.Execute();

            foreach (var sheet in templateSpreadSheet.Sheets)
            {
                var sourceSheetRange = $"{sheet.Properties.Title}!A:Z";
                var sourceValues = Service.Spreadsheets.Values.Get(templateWorksheetId, sourceSheetRange).Execute().Values;

                var targetSpreadsheetId = createdSpreadsheet.SpreadsheetId;
                var addSheetRequest = new AddSheetRequest { Properties = new SheetProperties { Title = sheet.Properties.Title } };
                var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest { Requests = new List<Request> { new Request { AddSheet = addSheetRequest } } };
                var response = Service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, targetSpreadsheetId).Execute();
                var newSheetId = response.Replies[0].AddSheet.Properties.Title;

                var targetSheetRange = $"{newSheetId}!A:Z";
                var valueRange = new ValueRange { Values = sourceValues };
                var updateRequest = Service.Spreadsheets.Values.Update(valueRange, targetSpreadsheetId, targetSheetRange);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                updateRequest.Execute();
            }

            DeleteSheet(createdSpreadsheet, 0);

            return createdSpreadsheet;
        }
    }

    private void DeleteSheet(Spreadsheet createdSpreadsheet, int sheetId)
    {
        var deleterequest = new Request
        {
            DeleteSheet = new DeleteSheetRequest
            {
                SheetId = sheetId
            }
        };

        var deleteSheet = new BatchUpdateSpreadsheetRequest
        {
            Requests = new List<Request> { deleterequest }
        };

        Service.Spreadsheets.BatchUpdate(deleteSheet, createdSpreadsheet.SpreadsheetId).Execute();
    }
}