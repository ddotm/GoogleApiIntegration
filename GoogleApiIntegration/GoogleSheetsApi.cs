using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace GoogleApiIntegration
{
	public class GoogleSheetsApi
	{
		private readonly SheetsService _sheetsService;

		public GoogleSheetsApi(string appName, ServiceAccountCredential serviceAccountCredential)
		{
			_sheetsService = new SheetsService(new BaseClientService.Initializer
			{
				ApplicationName = appName,
				HttpClientInitializer = serviceAccountCredential
			});
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="title"></param>
		/// <param name="sheetTitle"></param>
		/// <returns></returns>
		public async Task<Spreadsheet> CreateAsync(string title, string sheetTitle)
		{
			var requestBody = new Spreadsheet
			{
				Properties = new SpreadsheetProperties
				{
					Title = title
				},
				Sheets = new List<Sheet>
				{
					new Sheet
					{
						Properties = new SheetProperties
						{
							Title = sheetTitle,
							TabColor = new Color
							{
								Red = 0.9F,
								Green = 0.1F,
								Blue = 0.1F,
								Alpha = 0.7F
							},
							GridProperties = new GridProperties
							{
								ColumnCount = 50
							}
						}
					}
				}
			};

			SpreadsheetsResource.CreateRequest createRequest = _sheetsService.Spreadsheets.Create(requestBody);

			Spreadsheet createResponse = await createRequest.ExecuteAsync();
			return createResponse;
		}

		public async Task<BatchUpdateSpreadsheetResponse> RenameSheet(string spreadsheetId, int sheetId, string oldName, string newName)
		{
			var req = new BatchUpdateSpreadsheetRequest
			{
				Requests = new List<Request>
				{
					new Request
					{
						UpdateSheetProperties = new UpdateSheetPropertiesRequest
						{
							Properties = new SheetProperties
							{
								SheetId = sheetId,
								Title = newName
							}
						}
					}
				}
			};
			SpreadsheetsResource.BatchUpdateRequest request = new SpreadsheetsResource.BatchUpdateRequest(_sheetsService, req, spreadsheetId);
			var response = await request.ExecuteAsync();

			return response;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="spreadsheetId"></param>
		/// <param name="sheetTitle"></param>
		/// <param name="headers"></param>
		/// <param name="data"></param>
		/// <returns></returns>
		public async Task<BatchUpdateValuesResponse> BatchUpdateAsync(string spreadsheetId, string sheetTitle, List<object> headers, List<IList<object>> data)
		{
			// How the input data should be interpreted.
			const string valueInputOption = "USER_ENTERED";

			// The new values to apply to the spreadsheet.
			var reqData = new List<ValueRange>
			{
				new ValueRange
				{
					Range = $"{sheetTitle}!A1:C1",
					Values = new List<IList<object>>
					{
						headers
					}
				},
				new ValueRange
				{
					Range = $"{sheetTitle}!A2:C4",
					Values = data
				}
			};

			var requestBody = new BatchUpdateValuesRequest
			{
				ValueInputOption = valueInputOption,
				IncludeValuesInResponse = true,
				Data = reqData
			};

			SpreadsheetsResource.ValuesResource.BatchUpdateRequest batchUpdateRequest = _sheetsService.Spreadsheets.Values.BatchUpdate(requestBody, spreadsheetId);

			BatchUpdateValuesResponse batchUpdateResponse = await batchUpdateRequest.ExecuteAsync();

			return batchUpdateResponse;
		}
	}
}
