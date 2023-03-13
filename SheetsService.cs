using ConsoleToGoogleSheets.Mapers;
using ConsoleToGoogleSheets.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleToGoogleSheets
{
    public class SheetsService
    {
        static readonly string[] Scopes = { Google.Apis.Sheets.v4.SheetsService.Scope.Spreadsheets };
        static string ApplicationName;
        static string SpreadSheetId;
        static string SheetName;
        static string SecureAnswer;

        static Google.Apis.Sheets.v4.SheetsService service;

        public SheetsService(string applicationName, string spreadSheetId, string sheetName, string secureAnswer)
        {
            ApplicationName = applicationName;
            SpreadSheetId = spreadSheetId;
            SheetName = sheetName;
            SecureAnswer = secureAnswer;

            GoogleCredential credential;
            using (var stream = new FileStream(SecureAnswer, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }
            service = new Google.Apis.Sheets.v4.SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }
        public void ReadEntrys()
        {
            var range = $"{SheetName}!A1:F24002";
            var reguest = service.Spreadsheets.Values.Get(SpreadSheetId, range);

            var response = reguest.Execute();
            var values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    Console.WriteLine("{0} {1} {2} {3}", row[0], row[1], row[2], row[3] );
                }
            }
            else
                Console.WriteLine("Not fiend data");
        }
        public void CreateEntry()
        {
            var range = $"{SheetName}!A:F";
            var valueRange = new ValueRange();

            var objectList = new List<object>() { "1000", "new", "Entry" };
            valueRange.Values = new List<IList<object>> { objectList };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadSheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var response = appendRequest.Execute();
        }

        public List<Item> ReadItemsAsync()
        {
            var range = $"{SheetName}!A:N";
            var request = service.Spreadsheets.Values.Get(SpreadSheetId, range);
            var response = request.Execute();
            var values = response.Values;
            var items = ItemsMapper.MapFromRangeData(values);
            return items;
        }
        public async Task<AppendValuesResponse> AddItemsAsync(Item[] items )
        {
            var range = $"{SheetName}!A:N";
            var valueRange = ItemsMapper.MapToRangeData(items);
            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadSheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            return await appendRequest.ExecuteAsync().ConfigureAwait(false);
            //return await Task.Run(async () =>
            //{
            //    await Task.Delay(10000).ConfigureAwait(false);
            //    return await appendRequest.ExecuteAsync().ConfigureAwait(false);
            //}); 
           
        }
    }
}
