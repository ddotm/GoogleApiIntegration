using Google.Apis.Drive.v3.Data;
using GoogleApiIntegration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Reflection;
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
			const string spreadsheetTitle = "Important exported info";
			const string sheetTitle = "You precious data";

			var credential = GoogleCredentialManager.CreateGoogleCredential(serviceAcctEmail, configFile);

			var sheetsApi = new GoogleSheetsApi(appName, credential);
			var createResponse = await sheetsApi.CreateAsync(spreadsheetTitle, sheetTitle);
			Console.WriteLine("Create sheet response");
			Console.WriteLine(JsonConvert.SerializeObject(createResponse));
			var headers = GetHeaders();
			var data = GetData();
			var batchUpdateValuesResponse = await sheetsApi.BatchUpdateAsync(createResponse.SpreadsheetId, sheetTitle, headers, data);
			Console.WriteLine("Batch update sheet response");
			Console.WriteLine(JsonConvert.SerializeObject(batchUpdateValuesResponse));

			var driveApi = new GoogleDriveApi(appName, credential);
			var fileList = await driveApi.ListFilesAsync();
			Console.WriteLine("File list:");
			Console.WriteLine(JsonConvert.SerializeObject(fileList));
			var permission = await driveApi.ShareAsync(createResponse.SpreadsheetId, "dmitri.mogilevski@paradigmagency.com");
			// permission = await driveApi.ShareAsync(createResponse.SpreadsheetId, "jackporter05@gmail.com");
			// permission = await driveApi.ShareAsync(createResponse.SpreadsheetId,  "timothyjohneck@gmail.com");

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

		private static List<IList<object>> GetHeaders()
		{
			return new List<IList<object>>
			{
				new List<object>
				{
					"First name",
					"Last name",
					"Email",
					"Note"
				}
			};
		}

		private static List<string> GetColNames()
		{
			return new List<string>
			{
				"FirstName",
				"LastName",
				"Email",
				"Note"
			};
		}

		private static object GetPropValue(object src, string propName)
		{
			return src?.GetType().GetProperty(propName, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance)?.GetValue(src, null);
		}

		private static List<IList<object>> GetData()
		{
			var data = new List<IList<object>>();

			var sourceDataItems = new List<DataClass>
			{
				new DataClass
				{
					FirstName = "Dmitri",
					LastName = "Mogilevski",
					Email = "dmogilevski@gmail.com",
					Note = "Dmitri's note"
				},
				new DataClass
				{
					FirstName = "Jack",
					LastName = "Porter",
					Email = "jporter@someemail.com",
					Note = "I just wanted to jot this down"
				},
				new DataClass
				{
					FirstName = "Tim",
					LastName = "Eck",
					Email = "teck@someemail.com",
					Note = "Eck note"
				}
			};
			var colNames = GetColNames();

			foreach (var sourceDataItem in sourceDataItems)
			{
				var dataItem = new List<object>();
				foreach (var colName in colNames)
				{
					var val = GetPropValue(sourceDataItem, colName);
					dataItem.Add(val);
				}

				data.Add(dataItem);
			}

			return data;
		}
	}
}
