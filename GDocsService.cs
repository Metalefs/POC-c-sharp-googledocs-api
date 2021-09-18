using Google.Apis.Auth.OAuth2;
using Google.Apis.Docs.v1;
using Google.Apis.Docs.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Collections.Generic;
using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;
using ADE.Dominio.Models;

namespace GoogleDocsIntegration
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/docs.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { DocsService.Scope.DocumentsReadonly, DocsService.Scope.Documents };
        static string serviceAccountEmail = "ade-service@assistente-de-estagio-326401.iam.gserviceaccount.com";
        static string serviceAccountCredentialFilePath = "./key.p12";
        public static DocsService DocsService;
        static void Main(string[] args)
        {
            AuthenticateServiceAccount(serviceAccountEmail, serviceAccountCredentialFilePath, Scopes);

            // Define request parameters.
            String documentId = "15I3l7P-luymw-89zXcgcxo6zFYovyZ_jUh-tKxe6E4c";
            DocumentsResource.GetRequest request = DocsService.Documents.Get(documentId);

            // Prints the title of the requested doc:
            // https://docs.google.com/document/d/195j9eDD3ccgjQRttHhJPymLJUCOUjs-jmwTrekvdjFE/edit
            Document doc = request.Execute();
            Console.WriteLine("The title of the doc is: {0}", doc.Title);

        }

        /// <summary>
        /// Authenticating to Google using a Service account
        /// Documentation: https://developers.google.com/accounts/docs/OAuth2#serviceaccount
        /// </summary>
        /// <param name="serviceAccountEmail">From Google Developer console https://console.developers.google.com</param>
        /// <param name="serviceAccountCredentialFilePath">Location of the .p12 or Json Service account key file downloaded from Google Developer console https://console.developers.google.com</param>
        /// <returns>AnalyticsService used to make requests against the Analytics API</returns>
        static void AuthenticateServiceAccount(string serviceAccountEmail, string serviceAccountCredentialFilePath, string[] scopes)
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
                        ApplicationName = "Drive Authentication Sample",
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

        public async Task<Document> GetDocumentByID(string documentId)
        {
            var document = await DocsService.Documents.Get(documentId).ExecuteAsync();
            return document;
        }

        public async Task<Document> MergeTemplate(string documentId, List<Template> templateValues)
        {
            List<Request> requests = new List<Request>();
            var document = await DocsService.Documents.Get(documentId).ExecuteAsync();
            DocsService.Documents.Create(document);

            foreach(var template in templateValues)
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

            var result = DocsService.Documents.BatchUpdate(body, document.DocumentId).Execute();
            return await DocsService.Documents.Get(result.DocumentId).ExecuteAsync();
        }
    }
}