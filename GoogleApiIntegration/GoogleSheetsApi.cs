using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;

namespace GoogleApiIntegration
{
	public class GoogleSheetsApi
	{
		private readonly SheetsService _sheetsService;

		public GoogleSheetsApi(string appName, ServiceAccountCredential serviceAccountCredential)
		{
			_sheetsService = new SheetsService(new BaseClientService.Initializer
			{
				ApplicationName = "TempoApi",
				HttpClientInitializer = serviceAccountCredential
			});
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="title"></param>
		/// <returns></returns>
		public async Task<Spreadsheet> CreateAsync(string title)
		{
			var requestBody = new Spreadsheet
			{
				Properties = new SpreadsheetProperties
				{
					Title = "Jack's Great Spreadsheet"
				}
			};

			SpreadsheetsResource.CreateRequest createRequest = _sheetsService.Spreadsheets.Create(requestBody);

			Spreadsheet createResponse = await createRequest.ExecuteAsync();
			Console.WriteLine("Create sheet response");

			Console.WriteLine(JsonConvert.SerializeObject(createResponse));
			return createResponse;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="spreadsheetId"></param>
		/// <returns></returns>
		public async Task<BatchUpdateValuesResponse> BatchUpdateAsync(string spreadsheetId)
		{
			// How the input data should be interpreted.
			const string valueInputOption = "USER_ENTERED";

			// The new values to apply to the spreadsheet.
			var data = new List<ValueRange>
			{
				new ValueRange
				{
					Range = "Sheet1!A1:C1",
					Values = new List<IList<object>>
					{
						new List<object>
						{
							"First name",
							"Last name",
							"Email"
						}
					}
				},
				new ValueRange
				{
					Range = "Sheet1!A2:C2",
					Values = new List<IList<object>>
					{
						new List<object>
						{
							"Jack",
							"Porter",
							"someemail@something.com"
						}
					}
				},
				new ValueRange
				{
					Range = "Sheet1!A3:C3",
					Values = new List<IList<object>>
					{
						new List<object>
						{
							"Tim",
							"Eck",
							"teck@something.com"
						}
					}
				}
			};

			var requestBody = new BatchUpdateValuesRequest
			{
				ValueInputOption = valueInputOption,
				IncludeValuesInResponse = true,
				Data = data
			};

			SpreadsheetsResource.ValuesResource.BatchUpdateRequest batchUpdateRequest = _sheetsService.Spreadsheets.Values.BatchUpdate(requestBody, spreadsheetId);

			BatchUpdateValuesResponse batchUpdateResponse = await batchUpdateRequest.ExecuteAsync();
			Console.WriteLine("Batch update sheet response");
			Console.WriteLine(JsonConvert.SerializeObject(batchUpdateResponse));

			return batchUpdateResponse;
		}
	}
}
