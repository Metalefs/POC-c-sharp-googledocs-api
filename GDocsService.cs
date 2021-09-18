using Google.Apis.Auth.OAuth2;
using Google.Apis.Docs.v1;
using Google.Apis.Docs.v1.Data;
using Google.Apis.Services;
using System.Collections.Generic;
using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using ADE.Dominio.Models;
using Google.Apis.Drive.v3;

namespace GoogleDocsIntegration
{
    public class GDocsService
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/docs.googleapis.com-dotnet-quickstart.json
        public static string[] scopes = { DocsService.Scope.Documents, DocsService.Scope.Drive, DriveService.Scope.Drive };
        public static string serviceAccountEmail = "ade-service@assistente-de-estagio-326401.iam.gserviceaccount.com";
        public static string serviceAccountCredentialFilePath = "./key.p12";
        public static DocsService DocsService;
        public static DriveService DriveService;
       
        /// <summary>
        /// Authenticating to Google using a Service account
        /// Documentation: https://developers.google.com/accounts/docs/OAuth2#serviceaccount
        /// </summary>
        /// <param name="serviceAccountEmail">From Google Developer console https://console.developers.google.com</param>
        /// <param name="serviceAccountCredentialFilePath">Location of the .p12 or Json Service account key file downloaded from Google Developer console https://console.developers.google.com</param>
        /// <returns>AnalyticsService used to make requests against the Analytics API</returns>
        public void AuthenticateServiceAccount()
        {
            try
            {
                if (string.IsNullOrEmpty(serviceAccountCredentialFilePath))
                    throw new Exception("Path to the service account credentials file is required.");
                if (!File.Exists(serviceAccountCredentialFilePath))
                    throw new Exception("The service account credentials file does not exist at: " + serviceAccountCredentialFilePath);
                if (string.IsNullOrEmpty(serviceAccountEmail))
                    throw new Exception("ServiceAccountEmail is required.");

                // For Json file
                if (Path.GetExtension(serviceAccountCredentialFilePath).ToLower() == ".json")
                {
                    GoogleCredential credential;
                    using (var stream = new FileStream(serviceAccountCredentialFilePath, FileMode.Open, FileAccess.Read))
                    {
                        credential = GoogleCredential.FromStream(stream)
                             .CreateScoped(scopes);
                    }

                    // Create the  Analytics service.
                    DocsService = new DocsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = "Docs Service account Authentication Sample",
                    });
                    DriveService = new DriveService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = "Drive Service account Authentication Sample",
                    });
                }
                else if (Path.GetExtension(serviceAccountCredentialFilePath).ToLower() == ".p12")
                {   // If its a P12 file

                    var certificate = new X509Certificate2(serviceAccountCredentialFilePath, "notasecret", X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.Exportable);
                    var credential = new ServiceAccountCredential(new ServiceAccountCredential.Initializer(serviceAccountEmail)
                    {
                        Scopes = scopes
                    }.FromCertificate(certificate));

                    // Create the  Drive service.
                    DocsService = new DocsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = "Docs Authentication Sample",
                    }); 
                    DriveService = new DriveService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = "Drive Service account Authentication Sample",
                    });
                }
                else
                {
                    throw new Exception("Unsupported Service accounts credentials.");
                }

            }
            catch (Exception ex)
            {
                throw new Exception("CreateServiceAccountDriveFailed", ex);
            }
        }

        public static Document GetDocumentByID(string documentId)
        {
            var document = DocsService.Documents.Get(documentId).Execute();
            return document;
        }

        public async static Task<Document> GetDocumentByIDAsync(string documentId)
        {
            var document = await DocsService.Documents.Get(documentId).ExecuteAsync();
            return document;
        }

        public async static Task<Document> MergeTemplate(string documentId, List<Template> templateValues)
        {
            List<Request> requests = new List<Request>();
            var document = await DocsService.Documents.Get(documentId).ExecuteAsync();

            Google.Apis.Drive.v3.Data.File targetFile = new Google.Apis.Drive.v3.Data.File();
            //This will be the body of the request so probably you would want to modify this
            targetFile.Name = document.Title;
            targetFile.CopyRequiresWriterPermission = false;
            
            var copyrequest = DriveService.Files.Get(documentId);
            copyrequest.SupportsAllDrives = true;
            
            var cp_file = await copyrequest.ExecuteAsync();
            
            var copy_result = await DriveService.Files.Copy(targetFile, cp_file.Id).ExecuteAsync();

            foreach (var template in templateValues)
            {
                var repl = new Request();
                var substrMatchCriteria = new SubstringMatchCriteria();
                var replaceAlltext = new ReplaceAllTextRequest();

                substrMatchCriteria.Text = template.Key;
                replaceAlltext.ReplaceText = template.Value;

                replaceAlltext.ContainsText = substrMatchCriteria;
                repl.ReplaceAllText = replaceAlltext;

                requests.Add(repl);
            }

            BatchUpdateDocumentRequest body = new BatchUpdateDocumentRequest { Requests = requests };

            var result = DocsService.Documents.BatchUpdate(body, copy_result.Id).Execute();
            return await DocsService.Documents.Get(result.DocumentId).ExecuteAsync();
        }

        public Stream ExportPDF(string documentId)
        {
            var exported = DriveService.Files.Export(documentId, "application/pdf");

            Stream stream = new MemoryStream();
            exported.Download(stream);
            return stream;
        }

        public async Task<Stream> ExportPDFAsync(string documentId)
        {
            var exported = DriveService.Files.Export(documentId, "application/pdf");

            Stream stream = new MemoryStream();
            await exported.DownloadAsync(stream);
            return stream;
        }
    }
}