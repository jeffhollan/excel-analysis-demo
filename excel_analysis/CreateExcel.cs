using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using OfficeOpenXml;

namespace excel_analysis
{
    public static class CreateExcel
    {
        [FunctionName("CreateExcel")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");
            var neededSheets = req.Headers.GetValues("x-sheets").First();

            using (var package = new ExcelPackage(await req.Content.ReadAsStreamAsync()))
            {
                foreach (var sheet in package.Workbook.Worksheets) {
                    if (!neededSheets.Contains(sheet.Name))  {
                        package.Workbook.Worksheets.Delete(sheet.Name);
                    }
                }

                MemoryStream _stream = new MemoryStream();
                package.SaveAs(_stream);
                var res = req.CreateResponse();
                res.Content = new ByteArrayContent(_stream.ToArray());
                res.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                return res;
            }

            
        }
    }
}