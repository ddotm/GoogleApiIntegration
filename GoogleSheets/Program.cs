﻿using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Http;

namespace GoogleSheets
{
	internal class Program
	{
		private static async Task Main(string[] args)
		{
			ServiceAccountCredential credential = CreateGoogleCredential();

			var sheetsService = new SheetsService(new BaseClientService.Initializer
			{
				ApplicationName = "TempoApi",
				HttpClientInitializer = credential
			});

			Spreadsheet createResponse = await CreateAsync(sheetsService);

			BatchUpdateValuesResponse batchUpdateValuesResponse = await BatchUpdateAsync(sheetsService, createResponse.SpreadsheetId);

			var fileList = await ListFilesAsync(credential);

			await ShareAsync(credential, createResponse.SpreadsheetId, "dmitri.mogilevski@paradigmagency.com");

			Console.ReadKey();
		}

		/// <summary>
		/// Google Drive API -- List Files
		/// </summary>
		/// <param name="credential"></param>
		/// <returns></returns>
		private static async Task<IList<Google.Apis.Drive.v3.Data.File>> ListFilesAsync(ServiceAccountCredential credential)
		{
			var service = new DriveService(new BaseClientService.Initializer
			{
				HttpClientInitializer = credential,
				ApplicationName = "TempoApi"
			});

			FilesResource.ListRequest listRequest = service.Files.List();
			listRequest.PageSize = 10;
			listRequest.Fields = "nextPageToken, files(id, name)";

			// List files.
			IList<Google.Apis.Drive.v3.Data.File> files = (await listRequest.ExecuteAsync())
				.Files;

			Console.WriteLine("File list:");
			Console.WriteLine(JsonConvert.SerializeObject(files));

			return files;
		}

		private static async Task<Permission> ShareAsync(ServiceAccountCredential credential, string fileId, string email)
		{
			var service = new DriveService(new BaseClientService.Initializer
			{
				HttpClientInitializer = credential,
				ApplicationName = "TempoApi"
			});

			var permission = new Permission
			{
				Kind = "drive#permission",
				// The type of the grantee.Valid values are: -user - group - domain - anyone
				Type = "user",
				// currently allowed: - owner - organizer - fileOrganizer - writer - commenter - reader
				Role = "writer",
				EmailAddress = "dmitri.mogilevski@paradigmagency.com",
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

			var request = service.Permissions.Create(permission, fileId);
			permission = await request.ExecuteAsync();

			return permission;
		}

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


		private static async Task<BatchUpdateValuesResponse> BatchUpdateAsync(SheetsService sheetsService, string spreadsheetId)
		{
			// How the input data should be interpreted.
			const string valueInputOption = "USER_ENTERED";

			// The new values to apply to the spreadsheet.
			var data = new List<ValueRange>
			{
				new ValueRange
				{
					Range = "Sheet1!A1:B1",
					Values = new List<IList<object>>
					{
						new List<object>
						{
							"First name",
							"Last name"
						}
					}
				},
				new ValueRange
				{
					Range = "Sheet1!A2:B2",
					Values = new List<IList<object>>
					{
						new List<object>
						{
							"Jack",
							"Porter"
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

		private static ServiceAccountCredential CreateGoogleCredential()
		{
			ServiceAccountCredential credential;
			string[] scopes =
			{
				SheetsService.Scope.Spreadsheets,
				SheetsService.Scope.Drive,
				SheetsService.Scope.DriveFile
			};
			var serviceAccountEmail = "tempo-api@tempo-265222.iam.gserviceaccount.com";
			var jsonFile = "tempo-265222-88dada7b17c3.json";
			using (Stream stream = new FileStream(@jsonFile, FileMode.Open, FileAccess.Read, FileShare.Read))
			{
				credential = (ServiceAccountCredential)
					GoogleCredential.FromStream(stream).UnderlyingCredential;

				var initializer = new ServiceAccountCredential.Initializer(credential.Id)
				{
					User = serviceAccountEmail,
					Key = credential.Key,
					Scopes = scopes
				};
				credential = new ServiceAccountCredential(initializer);
			}

			return credential;
		}
	}
}
