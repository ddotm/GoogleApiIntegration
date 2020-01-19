using Google.Apis.Drive.v3.Data;
using GoogleApiIntegration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace GoogleSheets
{
	internal class Program
	{
		private static async Task Main(string[] args)
		{
			const string serviceAcctEmail = "tempo-api@tempo-265222.iam.gserviceaccount.com";
			const string configFile = "tempo-265222-88dada7b17c3.json";
			const string appName = "TempoApi";
			var credential = GoogleCredentialManager.CreateGoogleCredential(serviceAcctEmail, configFile);

			var sheetsApi = new GoogleSheetsApi(appName, credential);
			var createResponse = await sheetsApi.CreateAsync("Test spreadsheet");
			Console.WriteLine("Create sheet response");
			Console.WriteLine(JsonConvert.SerializeObject(createResponse));
			var batchUpdateValuesResponse = await sheetsApi.BatchUpdateAsync(createResponse.SpreadsheetId);
			Console.WriteLine("Batch update sheet response");
			Console.WriteLine(JsonConvert.SerializeObject(batchUpdateValuesResponse));

			var driveApi = new GoogleDriveApi(appName, credential);
			var fileList = await driveApi.ListFilesAsync();
			Console.WriteLine("File list:");
			Console.WriteLine(JsonConvert.SerializeObject(fileList));
			var permission = await driveApi.ShareAsync(createResponse.SpreadsheetId, "dmitri.mogilevski@paradigmagency.com");
			// permission = await ShareAsync(driveService, createResponse.SpreadsheetId, "jackporter05@gmail.com");

			// await DeleteAllExceptAsync(fileList, createResponse.SpreadsheetId, driveApi);

			Console.ReadKey();
		}

		private static async Task DeleteAllExceptAsync(IList<File> fileList, string fileId, GoogleDriveApi driveApi)
		{
			foreach (var file in fileList)
			{
				if (file.Id == fileId) continue;
				var resp = await driveApi.DeleteAsync(file.Id);
				Console.WriteLine($"Delete file response (file id/name {file.Id} {file.Name}): {resp}");
			}

			string emptyTrashResponse = await driveApi.EmptyTrashAsync();
			Console.WriteLine($"Empty trash response: {emptyTrashResponse}");
		}
	}
}
