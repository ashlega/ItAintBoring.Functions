using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using RestSharp;


namespace PowerBI
{

    public class PowerBIExportResult
    {
        public string errorMessage = null;
        public byte[] fileContent = null;
        public string fileContentType = null;
    }
    public class PowerBIHelper
    {
        const int SLEEP_TIME = 2000;
        string appId = null;
        string secret = null;
        string domain = null;

        public PowerBIHelper(string appId, string secret, string domain)
        {
            this.appId = appId;
            this.secret = secret;
            this.domain = domain;
        }
        private async Task<string> AuthenticatePowerBI()
        {
            
            var credentials = new ClientCredential(appId, secret);
            string authorityUri = $"https://login.microsoftonline.com/{domain}/";
            var authContext = new AuthenticationContext(authorityUri);
            var token = await authContext.AcquireTokenAsync("https://analysis.windows.net/powerbi/api", credentials);
            return token.AccessToken;
        }

        public async Task<PowerBIExportResult> ExportToFile(string groupId, string reportId)
        {
            PowerBIExportResult result = new PowerBIExportResult();

            var accessToken = await AuthenticatePowerBI();

            //Start the export
            var client = new RestClient($"https://api.powerbi.com/v1.0/myorg/groups/{groupId}/reports/{reportId}/ExportTo");
            var request = new RestRequest(Method.POST);
            request.AddHeader("cache-control", "no-cache");
            request.AddHeader("authorization", "Bearer " + accessToken);
            request.AddHeader("content-type", "application/json");
            request.AddParameter("application/json", "\r\n{\r\n\t\"format\": \"DOCX\"\r\n}", ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);
            
            if (response.IsSuccessful)
            {
                //Now wait for the export
                dynamic responseObject = JsonConvert.DeserializeObject(response.Content);
                string exportId = null;
                exportId = responseObject?.id;
                client = new RestClient($"https://api.powerbi.com/v1.0/myorg/groups/{groupId}/reports/{reportId}/exports/{exportId}");
                request = new RestRequest(Method.GET);
                request.AddHeader("cache-control", "no-cache");
                request.AddHeader("authorization", "Bearer " + accessToken);
                string exportStatus = "";
                //Wait till the export is ready
                do
                {
                    response = client.Execute(request);
                    System.Threading.Thread.Sleep(SLEEP_TIME);
                    if (response.IsSuccessful)
                    {
                        responseObject = JsonConvert.DeserializeObject(response.Content);
                        exportStatus = responseObject?.status;
                    }
                    
                } while (response.IsSuccessful && exportStatus != "Failed" && exportStatus != "Succeeded");

                //And download exported file
                if(response.IsSuccessful && exportStatus == "Succeeded")
                {
                    string resourceLocation = responseObject?.resourceLocation;
                    client = new RestClient(resourceLocation);
                    request = new RestRequest(Method.GET);
                    request.AddHeader("cache-control", "no-cache");
                    request.AddHeader("authorization", "Bearer " + accessToken);
                    response = client.Execute(request);
                    if (response.IsSuccessful)
                    {
                        result.fileContent = response?.RawBytes;
                        result.fileContentType = response?.ContentType;
                    }
                    else
                    {
                        result.errorMessage = response.Content;
                    }
                }
                else
                {
                    result.errorMessage = response.Content;
                }

            }
            else
            {
                result.errorMessage = response.Content;
            }

            return result;
        }

    }
}
