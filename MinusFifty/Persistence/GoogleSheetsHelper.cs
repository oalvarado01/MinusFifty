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
using System.Text.RegularExpressions;

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

        public async Task<GridRange> GetGridRange(string range)
        {
            GridRange outputRange = new GridRange();
            Match match = Regex.Match(range, @"(^.+)!(.+):(.+$)");
            GroupCollection data;
            if (match.Success)
            {
                data = match.Groups;
            }
            else
            {
                return outputRange;
            }

            Spreadsheet spreadsheet = await _service.Spreadsheets.Get(SheetId()).ExecuteAsync();
            Sheet targetSheet = null;
            foreach (Sheet sheet in spreadsheet.Sheets)
            {
                if (sheet.Properties.Title.Equals(data[1].ToString()))
                {
                    targetSheet = sheet;
                    break;
                }
            }

            if (targetSheet == null)
            {
                return outputRange;
            }

            GroupCollection data1 = null;
            GroupCollection data2 = null;

            int startIdx = -1;
            int endIdx = -1;

            Match co1 = Regex.Match(data[2].ToString(), @"(\D+)(\d+)");
            if (co1.Success)
            {
                data1 = co1.Groups;
                int.TryParse(data1[2].ToString(), out startIdx);
                --startIdx;
            }
            Match co2 = Regex.Match(data[3].ToString(), @"(\D+)(\d+)");
            if (co2.Success)
            {
                data2 = co2.Groups;
                int.TryParse(data2[2].ToString(), out endIdx);
                --endIdx;
            }

            outputRange.SheetId = targetSheet.Properties.SheetId;
            if (startIdx >= 0)
            {
                outputRange.StartRowIndex = startIdx;
            }
            if (endIdx >= 0)
            {
                outputRange.EndRowIndex = endIdx;
            }
            outputRange.StartColumnIndex = data1 != null ? to10(data1[1].ToString()) : to10(data[2].ToString());
            outputRange.EndColumnIndex = data2 != null ? to10(data2[1].ToString(), 1) : to10(data[3].ToString());

            return outputRange;
        }

        private int to10(string a1, int add = 0)
        {
            int lvl = a1.Length - 1;
            int val = add + (int)Math.Pow(26, lvl) * (a1.ToUpper()[0] - 64 - (lvl > 0 ? 0 : 1));
            return lvl > 0 ? to10(a1.Substring(1), val) : val;
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
                if (row.First().ToString().ToLower().Contains(needle.ToLower()))
                    return i;
            }
            return -1;
        }

        public int SubIndexInRange(ValueRange haystack, string key, int subColumn, string needle)
        {
            for (int i = 0; i < haystack.Values.Count; ++i)
            {
                IList<object> row = haystack.Values[i];
                if (row.First().ToString().ToLower().Contains(key.ToLower()) && row.Count > subColumn && row[subColumn].ToString().Contains(needle))
                    return i;
            }
            return -1;
        }

        public async Task<BatchUpdateSpreadsheetResponse> DeleteRowsAsync(int tabId, int startRow, int endRow = -1)
        {
            if (endRow < 0)
                endRow = startRow + 1;

            Request request = new Request
            {
                DeleteDimension = new DeleteDimensionRequest
                {
                    Range = new DimensionRange
                    {
                        Dimension = "ROWS",
                        SheetId = tabId,
                        StartIndex = startRow,
                        EndIndex = endRow
                    }
                }
            };
            BatchUpdateSpreadsheetRequest batchRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>()
            };
            batchRequest.Requests.Add(request);
            return await _service.Spreadsheets.BatchUpdate(batchRequest, SheetId()).ExecuteAsync();
        }

        public async Task<BatchUpdateSpreadsheetResponse> TransactionAsync(IList<Request> requests)
        {
            BatchUpdateSpreadsheetRequest batchRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = requests
            };
            return await _service.Spreadsheets.BatchUpdate(batchRequest, SheetId()).ExecuteAsync();
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
