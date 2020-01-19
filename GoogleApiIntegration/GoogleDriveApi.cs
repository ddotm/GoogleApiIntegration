using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace GoogleApiIntegration
{
	public class GoogleDriveApi
	{
		private readonly DriveService _driveService;

		public GoogleDriveApi(string appName, ServiceAccountCredential serviceAccountCredential)
		{
			_driveService = new DriveService(new BaseClientService.Initializer
			{
				ApplicationName = appName,
				HttpClientInitializer = serviceAccountCredential
			});
		}

		/// <summary>
		/// Google Drive API -- List Files
		/// </summary>
		/// <returns></returns>
		public async Task<IList<File>> ListFilesAsync()
		{
			FilesResource.ListRequest listRequest = _driveService.Files.List();
			listRequest.PageSize = 50;
			listRequest.Fields = "nextPageToken, files(id, name)";

			// List files.
			IList<File> files = (await listRequest.ExecuteAsync())
				.Files;
			return files;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="fileId"></param>
		/// <param name="email"></param>
		/// <returns></returns>
		public async Task<Permission> ShareAsync(string fileId, string email)
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

			var request = _driveService.Permissions.Create(permission, fileId);
			permission = await request.ExecuteAsync();

			return permission;
		}
	}
}
