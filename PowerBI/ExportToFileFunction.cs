using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace PowerBI
{
    public static class ExportToFileFunction
    {
        static string appid = null; 
        static string secret = null;
        static string domain = null;


        [FunctionName("ExportToFile")]
        public static async Task<IActionResult> ExportToFile(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            string error = null;
            
            secret = Environment.GetEnvironmentVariable("SECRET_KEY", EnvironmentVariableTarget.Process);
            appid = Environment.GetEnvironmentVariable("CLIENTID_KEY", EnvironmentVariableTarget.Process);
            domain = Environment.GetEnvironmentVariable("DOMAIN", EnvironmentVariableTarget.Process);

            if (appid == null)  error = "CLIENTID_KEY missing from the configuration";
            if (secret == null) error = "SECRET_KEY is missing from the configuration";
            if (domain == null) error = "DOMAIN is missing from the configuration";

            if (error == null)
            {
                string groupId = req.Query["groupId"];
                string reportId = req.Query["reportId"];

                string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
                dynamic data = JsonConvert.DeserializeObject(requestBody);
                groupId = groupId ?? data?.groupId;
                reportId = reportId ?? data?.reportId;

                PowerBIHelper pbi = new PowerBIHelper(appid, secret, domain);
                var response = await pbi.ExportToFile(groupId, reportId);
                if (response.errorMessage == null)
                {
                    var fileResult = new FileContentResult(response.fileContent, response.fileContentType);
                    return fileResult;
                }
                else
                {
                    error = response.errorMessage;
                    
                }
            }
            var contentResult = new ContentResult();
            contentResult.StatusCode = 500;
            contentResult.Content = error;
            return contentResult;
        }
    }
}
