using Google.Apis.Auth.OAuth2;

using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace SheetsQuickstart
{
    // Class to demonstrate the use of Sheets list values API
    class Program
    {
        /* Global instance of the scopes required by this quickstart.
         If modifying these scopes, delete your previously saved token.json/ folder. */
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart";
        static readonly string sheet = "Sheet1";

        static void Main(string[] args)
        {
            try
            {
                GoogleCredential credential;
                //UserCredential credential;
                // Load client secrets.
                using (var stream =
                       new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    /* The file token.json stores the user's access and refresh tokens, and is created
                     automatically when the authorization flow completes for the first time. */
                    credential = GoogleCredential.FromStream(stream);
                    credential = credential.CreateScoped(new string[] { SheetsService.Scope.Spreadsheets });
                    
                    
                }

                // Create Google Sheets API service.
                var service = new SheetsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName
                });
                

                // Define request parameters.
                String spreadsheetId = "1XAv_RbC3ZapltCZgNrthIxm4FJKMBg0DrbF8SneElBs";
                var range = $"{sheet}!A:H";
               
                var valueRange = new ValueRange();

                var objectList = new List<object>() { "Hello!", "This", "was", "inserted", "via", "C#", "", "" };
                var objectList1 = new List<object>() { "This", "is", "example", "of", ".NET", "google", "sheets", "API" };
                valueRange.Values = new List<IList<object>> { objectList, objectList1 };

                var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var appendResponse = appendRequest.Execute();

                SpreadsheetsResource.ValuesResource.GetRequest request =
                   service.Spreadsheets.Values.Get(spreadsheetId, range);
                // Prints the names and majors of students in a sample spreadsheet:
                // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit

                ValueRange response = request.Execute();
                IList<IList<Object>> values = response.Values;
                if (values == null || values.Count == 0)
                {
                    Console.WriteLine("No data found.");
                    return;
                }
                /*Console.WriteLine("Name, Major");
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 1.
                    Console.WriteLine("{0}, {1}", row[0], row[1]);
                }*/
            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}