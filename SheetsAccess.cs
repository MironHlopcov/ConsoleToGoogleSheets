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
    public class SheetsAccess
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "ConsoleToGoogleSheets";
        static readonly string SpreadsheetId = "1HwvBHPkJjDxHJgBFLM82WV51N2AVZDU339Qw-nc6FYk";
        static readonly string sheet = "Table-1";

        static SheetsService service;

        public SheetsAccess()
        {
            GoogleCredential credential;
            using (var stream = new FileStream("secure-answer.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }
            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }
        public void ReadEntrys()
        {
            var range = $"{sheet}!A1:F24002";
            var reguest = service.Spreadsheets.Values.Get(SpreadsheetId, range);

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
            var range = $"{sheet}!A:F";
            var valueRange = new ValueRange();

            var objectList = new List<object>() { "1000", "new", "Entry" };
            valueRange.Values = new List<IList<object>> { objectList };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var response = appendRequest.Execute();
        }

    }
}
