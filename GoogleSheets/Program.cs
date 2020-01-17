using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Http;

namespace GoogleSheets
{
	internal class Program
	{
		// If modifying these scopes, delete your previously saved credentials
		// at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
		static string[] Scopes = {SheetsService.Scope.SpreadsheetsReadonly};
		static string ApplicationName = "Google Sheets API .NET Quickstart";

		private static async Task Main(string[] args)
		{
			await CreateSheetAsync();
		}

		private static async Task CreateSheetAsync()
		{
			ServiceAccountCredential credential = createGoogleCredential();

			var sheetsService = new SheetsService(new BaseClientService.Initializer
			{
				ApplicationName = "TempoApi",
				HttpClientInitializer = credential
			});

			var requestBody = new Spreadsheet
			{
				Properties = new SpreadsheetProperties
				{
					Title = "Jack's Great Spreadsheet"
				}
			};

			SpreadsheetsResource.CreateRequest request = sheetsService.Spreadsheets.Create(requestBody);

			Spreadsheet response = await request.ExecuteAsync();

			Console.WriteLine(JsonConvert.SerializeObject(response));

			Console.ReadKey();
		}

		private static ServiceAccountCredential createGoogleCredential()
		{
			// GoogleCredential.FromFile("tempo-265222-88dada7b17c3.json");
			ServiceAccountCredential credential;
			string[] Scopes = {SheetsService.Scope.Spreadsheets};
			string serviceAccountEmail = "tempo-api@tempo-265222.iam.gserviceaccount.com";
			string jsonfile = "tempo-265222-88dada7b17c3.json";
			using (Stream stream = new FileStream(@jsonfile, FileMode.Open, FileAccess.Read, FileShare.Read))
			{
				credential = (ServiceAccountCredential)
					GoogleCredential.FromStream(stream).UnderlyingCredential;

				var initializer = new ServiceAccountCredential.Initializer(credential.Id)
				{
					User = serviceAccountEmail,
					Key = credential.Key,
					Scopes = Scopes
				};
				credential = new ServiceAccountCredential(initializer);
			}

			return credential;
		}
	}
}
