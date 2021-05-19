using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Data = Google.Apis.Sheets.v4.Data;

namespace Hangfire
{
    public class SheetsAdjusterAsync
    {
        UserCredential credential;
        public static string[] Scopes = { SheetsService.Scope.Drive };
        public static string ApplicationName = "YOUR APPLICATION NAME";
        public SheetsService _service = new SheetsService(new BaseClientService.Initializer());
        public string spreadsheetId = "MAIN_SPREADSHEET_ID";
        public string MainRange = "MAIN_RANGE";
        public List<string> ReadResults { get; set; }

        public void Authorize()
        {
            using (var stream =
               new FileStream("PATH_TO_YOUR_CREDENTIALS_JSON_FILE", FileMode.Open, FileAccess.Read))
            {

                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            _service = service;
            SpreadsheetsResource.ValuesResource.GetRequest request = _service.Spreadsheets.Values.Get(spreadsheetId, MainRange);
        }

        public List<string> ReadEntries(string range, string spreadsheetId)
        {
                var request = this._service.Spreadsheets.Values.Get(spreadsheetId, range);
                var response = request.Execute();
                var values = response.Values;
                if (values != null && values.Count > 0)
                {
                    List<string> newlist = new List<string>();
                    foreach (var row in values)
                    {
                        try
                        {
                            newlist.Add((string)row[0]);
                        }
                        catch (Exception ex)
                        {
                            newlist.Add(ex.Message);
                        }
                    }
                    return newlist;
                }
                return null;
        }

        public string ReadSingleCell(string range, string spreadsheetId)
        {
                var request = this._service.Spreadsheets.Values.Get(spreadsheetId, range);
                var response = request.Execute();
                var values = response.Values;
                if (values != null && values.Count > 0)
                {
                    return (string)values[0][0];
                }
                else
                {
                    return "no value";
                }
        }

        public async Task ReadEntriesAsync(string range, string spreadsheetId)
        {

                var request = this._service.Spreadsheets.Values.Get(spreadsheetId, range);
                var response = await request.ExecuteAsync();
                var values = response.Values;
                if (values != null && values.Count > 0)
                {
                    ReadResults.Clear();
                    foreach (var row in values)
                    {
                        try
                        {
                            ReadResults.Add((string)row[0]);
                        }
                        catch (Exception ex)
                        {
                            ReadResults.Add(ex.Message);
                        }
                    }
                }

        }

        public async Task UpdateEntries(List<object> gelenliste, string xrange, string dimension, string sheetid)
        {

                var valueRange = new ValueRange();
                valueRange.Values = new List<IList<object>> { gelenliste };
                valueRange.MajorDimension = dimension;
                var updateRequest = this._service.Spreadsheets.Values.Update(valueRange, sheetid, xrange);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var updateResponse = await updateRequest.ExecuteAsync();

        }

        public async Task CreateEntries(List<object> gelenliste, string xrange, string dimension, string sheetid)
        {

                var valueRange = new ValueRange();
                valueRange.Values = new List<IList<object>> { gelenliste };
                valueRange.MajorDimension = dimension;

                var appendRequest = this._service.Spreadsheets.Values.Append(valueRange, sheetid, xrange);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var appendResponse = await appendRequest.ExecuteAsync();

        }

        public async Task ClearArea(string range, string spreadsheetId)
        {
            Data.ClearValuesRequest requestBody = new Data.ClearValuesRequest();
            SpreadsheetsResource.ValuesResource.ClearRequest request = _service.Spreadsheets.Values.Clear(requestBody, spreadsheetId, range);
            Data.ClearValuesResponse response = await request.ExecuteAsync();
        }

        public async Task CreateSheet(string title, string ssid)
        {

            var addSheetRequest = new AddSheetRequest();
            addSheetRequest.Properties = new SheetProperties();
            addSheetRequest.Properties.Title = title;
            BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = new List<Request>();
            batchUpdateSpreadsheetRequest.Requests.Add(new Request { AddSheet = addSheetRequest });

            var batchUpdateRequest = await this._service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, ssid).ExecuteAsync();
        }
    }
}
