using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Net.Http;

namespace ItAintBoring.Functions
{
    public static class WordTemplates
    {
        [FunctionName("DocumentsMerge")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string name = req.Query["name"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            string firstFile = data?.firstFile;
            string secondFile = data?.secondFile;
            DocumentHandler dh = new DocumentHandler();
            byte[] byteFile1 = Convert.FromBase64String(firstFile);
            byte[] byteFile2 = Convert.FromBase64String(secondFile);
            MemoryStream streamFile1 = new MemoryStream();
            streamFile1.Write(byteFile1, 0, byteFile1.Length);
            MemoryStream streamFile2 = new MemoryStream(byteFile2, false);
            
            dh.Merge(streamFile1, streamFile2);

            
            FileContentResult result = new FileContentResult(streamFile1.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessing")
            {
                FileDownloadName = "result.docx"
            };
            return result;
        }

    }
}
