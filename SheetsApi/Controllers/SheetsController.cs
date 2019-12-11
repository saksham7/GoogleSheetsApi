using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Download;
using Google.Apis.Drive.v2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.AspNetCore.Mvc;

namespace SheetsApi.Controllers
{
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class SheetsController : ControllerBase
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        const string SpreadsheetId = "1pfMfqzTLGYXoxd_a2pY4P_-ucHmHrV30YNCnPU77YJs";
        const string GoogleCredentialsFileName = "google-credentials.json";


        [HttpPost]
        public async Task<UpdateValuesResponse> WriteData()
        {
            var serviceValues = GetSheetsService().Spreadsheets.Values;
            var data = new List<IList<object>>();
            data.Add(new List<object> { "test1", "Developer","22","MS","DEH", "India" });
            data.Add(new List<object> { "test2", "Developer", "25", "MS", "DEH", "India" });
            data.Add(new List<object> { "test3", "Developer", "30", "MS", "DEH", "India" });
            return await WriteAsync(serviceValues,"A",4,data);
        }

        [HttpGet]
        public async Task<IActionResult> DownloadSheet()
        {
            string[] Scope = { "https://www.googleapis.com/auth/drive.readonly" };
            using (var stream = new FileStream(GoogleCredentialsFileName, FileMode.Open, FileAccess.Read))
            {
                DriveService service = new DriveService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(Scope)
                });
                var request = service.Files.Export(SpreadsheetId, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

                request.MediaDownloader.ProgressChanged +=
                (IDownloadProgress progress) =>
                {
                    switch (progress.Status)
                    {
                        case DownloadStatus.Downloading:
                            {
                                Console.WriteLine(progress.BytesDownloaded);
                                break;
                            }
                        case DownloadStatus.Completed:
                            {
                                Console.WriteLine("Download complete.");
                                break;
                            }
                        case DownloadStatus.Failed:
                            {
                                Console.WriteLine("Download failed.");
                                break;
                            }
                    }
                };
                using (MemoryStream filestream = new MemoryStream())
                {
                    await request.DownloadAsync(filestream);
                    return File(filestream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "test");
                }
            }

        }

        private static SheetsService GetSheetsService()
        {
            using (var stream = new FileStream(GoogleCredentialsFileName, FileMode.Open, FileAccess.Read))
            {
                var serviceInitializer = new BaseClientService.Initializer
                {
                    HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
                };
                return new SheetsService(serviceInitializer);
            }
        }

        private static async Task<UpdateValuesResponse> WriteAsync(SpreadsheetsResource.ValuesResource valuesResource, string col, int row, List<IList<object>> dataSource)
        {

            var range = "Sheet1" + "!" + col + row; // "Basic!B111";
            var valueRange = new ValueRange { MajorDimension = "ROWS" };

            // var objectList = new List<object> { "abc","123" };
            valueRange.Values = dataSource;

            var update = valuesResource.Update(valueRange, SpreadsheetId, range);
            update.ValueInputOption =
                SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            return await update.ExecuteAsync();
        }
        private static async Task<string> ReadAsync(SpreadsheetsResource.ValuesResource valuesResource)
        {
            const string ReadRange = "Sheet1!A:B";
        
            var response = await valuesResource.Get(SpreadsheetId, ReadRange).ExecuteAsync();
            var values = response.Values;

            if (values == null || !values.Any())
            {
                // Console.WriteLine("No data found.");
                return "";
            }
            var header = string.Join(" ", values.First().Select(r => r.ToString()));
            Console.WriteLine($"Header: {header}");
            return header;
        }
    }
}
