using System.Collections.Generic;
using System.IO;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

namespace BootCamperMachineConfiguration.Helper
{
    public static class GoogleDriveHelper
    {
        #region Private Variables

        private const string ApplicationName = "Machine Specs for BootCampers";
        private const string MajorDimension = "COLUMNS";
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };

        #endregion

        public static ValueRange GetValuesFromGoogleDrive(string spreadSheetId, string range,
            SpreadsheetsResource.ValuesResource.GetRequest.MajorDimensionEnum majorDimensionEnum =
                SpreadsheetsResource.ValuesResource.GetRequest.MajorDimensionEnum.ROWS)
        {
            var request = GetSheetService().Spreadsheets.Values.Get(spreadSheetId, range);
            request.MajorDimension = majorDimensionEnum;
            return request.Execute();
        }

        public static void UpdateValuesInGoogleDrive(string spreadSheetId, string range, string value)
        {
            var valueRange = new ValueRange
            {
                MajorDimension = MajorDimension,
                Values = new List<IList<object>> { new List<object> { value } }
            };

            var update = GetSheetService().Spreadsheets.Values.Update(valueRange, spreadSheetId, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            update.Execute();
        }

        private static SheetsService GetSheetService()
        {
            UserCredential credential;
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                var credPath = "token.json";
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");
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
                ApplicationName = ApplicationName
            });
            return service;
        }
    }
}
