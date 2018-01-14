using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using JsonConfig;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;

namespace MinusFifty
{
    class GoogleSheetsHelper
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "MinusFifty";

        private SheetsService _service;
        public SheetsService Service
        {
            get
            {
                return _service;
            }
        }

        static GoogleSheetsHelper _instance;
        public static GoogleSheetsHelper Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new GoogleSheetsHelper();
                    _instance.Init();
                }
                return _instance;
            }
        }

        private void Init()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            _service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        private string SheetId()
        {
            return Config.Global.SheetId;
        }

        public ValueRange Get(string range)
        {
            return _service.Spreadsheets.Values.Get(SheetId(), range).Execute();
        }

        public async Task<ValueRange> GetAsync(string range)
        {
            return await _service.Spreadsheets.Values.Get(SheetId(), range).ExecuteAsync();
        }

        public int IndexInRange(ValueRange haystack, string needle)
        {
            for (int i = 0; i < haystack.Values.Count; ++i)
            {
                IList<object> row = haystack.Values[i];
                if (string.Compare(row.First().ToString(), needle, true) == 0)
                    return i;
            }
            return -1;
        }

        public UpdateValuesResponse Update(string range, ValueRange body)
        {
            UpdateRequest request = _service.Spreadsheets.Values.Update(body, SheetId(), range);
            request.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;
            return request.Execute();
        }

        public async Task<UpdateValuesResponse> UpdateAsync(string range, ValueRange body)
        {
            UpdateRequest request = _service.Spreadsheets.Values.Update(body, SheetId(), range);
            request.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;
            return await request.ExecuteAsync();
        }

        public AppendValuesResponse Append(string range, ValueRange body)
        {
            AppendRequest request = _service.Spreadsheets.Values.Append(body, SheetId(), range);
            request.ValueInputOption = AppendRequest.ValueInputOptionEnum.USERENTERED;
            return request.Execute();
        }

        public async Task<AppendValuesResponse> AppendAsync(string range, ValueRange body)
        {
            AppendRequest request = _service.Spreadsheets.Values.Append(body, SheetId(), range);
            request.ValueInputOption = AppendRequest.ValueInputOptionEnum.USERENTERED;
            return await request.ExecuteAsync();
        }

        public void TestSheets()
        {
            // Define request parameters.
            String spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
            String range = "Class Data!A2:E";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    Service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                Console.WriteLine("Name, Major");
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Console.WriteLine("{0}, {1}", row[0], row[4]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Console.Read();
        }
    }
}
