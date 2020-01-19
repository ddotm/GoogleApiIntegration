using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Http;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using GoogleApiIntegration;

namespace GoogleSheets
{
	internal class Program
	{
		private static async Task Main(string[] args)
		{
			ServiceAccountCredential credential = GoogleCredentialManager.CreateGoogleCredential("tempo-api@tempo-265222.iam.gserviceaccount.com", "tempo-265222-88dada7b17c3.json");

			var sheetsService = new SheetsService(new BaseClientService.Initializer
			{
				ApplicationName = "TempoApi",
				HttpClientInitializer = credential
			});

			Spreadsheet createResponse = await CreateAsync(sheetsService);

			BatchUpdateValuesResponse batchUpdateValuesResponse = await BatchUpdateAsync(sheetsService, createResponse.SpreadsheetId);

			var driveService = new DriveService(new BaseClientService.Initializer
			{
				HttpClientInitializer = credential,
				ApplicationName = "TempoApi"
			});

			var fileList = await ListFilesAsync(driveService);

			var permission = await ShareAsync(driveService, createResponse.SpreadsheetId, "dmitri.mogilevski@paradigmagency.com");
			permission = await ShareAsync(driveService, createResponse.SpreadsheetId, "jackporter05@gmail.com");

			Console.ReadKey();
		}

		/// <summary>
		/// Google Drive API -- List Files
		/// </summary>
		/// <param name="driveService"></param>
		/// <returns></returns>
		private static async Task<IList<Google.Apis.Drive.v3.Data.File>> ListFilesAsync(DriveService driveService)
		{
			FilesResource.ListRequest listRequest = driveService.Files.List();
			listRequest.PageSize = 50;
			listRequest.Fields = "nextPageToken, files(id, name)";

			// List files.
			IList<Google.Apis.Drive.v3.Data.File> files = (await listRequest.ExecuteAsync())
				.Files;

			Console.WriteLine("File list:");
			Console.WriteLine(JsonConvert.SerializeObject(files));

			return files;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="driveService"></param>
		/// <param name="fileId"></param>
		/// <param name="email"></param>
		/// <returns></returns>
		private static async Task<Permission> ShareAsync(DriveService driveService, string fileId, string email)
		{
			var permission = new Permission
			{
				Kind = "drive#permission",
				// The type of the grantee.Valid values are: -user - group - domain - anyone
				Type = "user",
				// currently allowed: - owner - organizer - fileOrganizer - writer - commenter - reader
				Role = "writer",
				EmailAddress = email,
				PermissionDetails = new List<Permission.PermissionDetailsData>
				{
					new Permission.PermissionDetailsData
					{
						// currently possible: -file - member
						PermissionType = "member",
						// currently possible: - organizer - fileOrganizer - writer - commenter - reader
						Role = "writer"
					}
				}
			};

			var request = driveService.Permissions.Create(permission, fileId);
			permission = await request.ExecuteAsync();

			return permission;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sheetsService"></param>
		/// <returns></returns>
		private static async Task<Spreadsheet> CreateAsync(SheetsService sheetsService)
		{
			var requestBody = new Spreadsheet
			{
				Properties = new SpreadsheetProperties
				{
					Title = "Jack's Great Spreadsheet"
				}
			};

			SpreadsheetsResource.CreateRequest createRequest = sheetsService.Spreadsheets.Create(requestBody);

			Spreadsheet createResponse = await createRequest.ExecuteAsync();
			Console.WriteLine("Create sheet response");

			Console.WriteLine(JsonConvert.SerializeObject(createResponse));
			return createResponse;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sheetsService"></param>
		/// <param name="spreadsheetId"></param>
		/// <returns></returns>
		private static async Task<BatchUpdateValuesResponse> BatchUpdateAsync(SheetsService sheetsService, string spreadsheetId)
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

			SpreadsheetsResource.ValuesResource.BatchUpdateRequest batchUpdateRequest = sheetsService.Spreadsheets.Values.BatchUpdate(requestBody, spreadsheetId);

			BatchUpdateValuesResponse batchUpdateResponse = await batchUpdateRequest.ExecuteAsync();
			Console.WriteLine("Batch update sheet response");
			Console.WriteLine(JsonConvert.SerializeObject(batchUpdateResponse));

			return batchUpdateResponse;
		}
	}
}
