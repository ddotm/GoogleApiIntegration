using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;
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
			ServiceAccountCredential credential = GoogleCredentialManager.CreateGoogleCredential("tempo-api@tempo-265222.iam.gserviceaccount.com", "tempo-265222-88dada7b17c3.json");
			
			var sheetsApi = new GoogleSheetsApi("TempoApi", credential);
			Spreadsheet createResponse = await sheetsApi.CreateAsync("Jack's Great Spreadsheet");
			BatchUpdateValuesResponse batchUpdateValuesResponse = await sheetsApi.BatchUpdateAsync(createResponse.SpreadsheetId);

			var driveService = new DriveService(new BaseClientService.Initializer
			{
				HttpClientInitializer = credential,
				ApplicationName = "TempoApi"
			});

			var fileList = await ListFilesAsync(driveService);

			var permission = await ShareAsync(driveService, createResponse.SpreadsheetId, "dmitri.mogilevski@paradigmagency.com");
			// permission = await ShareAsync(driveService, createResponse.SpreadsheetId, "jackporter05@gmail.com");

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
	}
}
