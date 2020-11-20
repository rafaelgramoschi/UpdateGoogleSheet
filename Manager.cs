using Google.Apis.Auth.OAuth2;
using Google.Apis.Script.v1;
using Google.Apis.Script.v1.Data;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using Nancy.Json;
using System.Collections;

namespace CreateGoogleSheet
{
    class Manager
    {

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "New Google Sheets using API in .NET";
        static string spreadSheetId = "1RwO8N12NkFO7kBES7CHdRbKXK8yNd6fOFnEjV_w_qVE";
        public static SheetsService service;
        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            popolaGSheet();
        }

        public static void popolaGSheet()
        {
            var data = GenerateData();
            UpdateData(data);
        }


        private static List<IList<Object>> GenerateData()
        {
            List<ArrayList> myArr = new List<ArrayList>()
            {
                new ArrayList(){ "TIPO LAVORO", "ORE LAVORATE", "ORE FATTURATE", "DIFFERENZA" },
                new ArrayList(){ "webdev", 8, 9, 1 },
                new ArrayList(){ "system", 1, 1, 0 },
                new ArrayList(){ "other ", 2, 0, -2 },
            };

            List<IList<Object>> objNewRecords = new List<IList<Object>>();
            
            for (int x = 0; x < myArr.Count; x++)
            {
                IList<Object> obj = new List<Object>();
                for (int y = 0; y < myArr[0].Count; y++)
                {
                    obj.Add(myArr[x][y]);
                }
                objNewRecords.Add(obj);
            }
            return objNewRecords;
        }

        public static string UpdateData(List<IList<object>> data)
        {
            String range = "Page1!A1";
            string valueInputOption = "USER_ENTERED";

            // The new values to apply to the spreadsheet.
            List<ValueRange> updateData = new List<ValueRange>();
            var dataValueRange = new ValueRange();
            dataValueRange.Range = range;
            dataValueRange.Values = data;
            updateData.Add(dataValueRange);

            BatchUpdateValuesRequest requestBody = new BatchUpdateValuesRequest();
            requestBody.ValueInputOption = valueInputOption;
            requestBody.Data = updateData;

            var request = service.Spreadsheets.Values.BatchUpdate(requestBody, spreadSheetId);

            BatchUpdateValuesResponse response = request.Execute();
            // Data.BatchUpdateValuesResponse response = await request.ExecuteAsync(); // For async 

            return JsonConvert.SerializeObject(response);
        }
    }

}