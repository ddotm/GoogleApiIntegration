using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using System.IO;

namespace GoogleApiIntegration
{
	public class GoogleCredentialManager
	{
		public static ServiceAccountCredential CreateGoogleCredential(string serviceAccountEmail, string configFileName)
		{
			ServiceAccountCredential credential;
			string[] scopes =
			{
				SheetsService.Scope.Spreadsheets,
				SheetsService.Scope.Drive,
				SheetsService.Scope.DriveFile
			};

			using (Stream stream = new FileStream(configFileName, FileMode.Open, FileAccess.Read, FileShare.Read))
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
