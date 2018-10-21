using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Custom;
using SixLabors.ImageSharp;

namespace FunctionApp1
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            Run();
            // parse query parameter

            return null;
        }

        private static void Run()
        {
            using (ExcelPackage templatePackage = new ExcelPackage(new System.IO.FileInfo(@"C:\Temp\template.xlsx")))
            {
                var image = Image.Load(@"C:\Temp\temp.jpg");
                using (MemoryStream ms = new MemoryStream())
                {
                    image.SaveAsJpeg(ms);
                    var sheet = templatePackage.Workbook.Worksheets[0];
                    var imageModel = new ImageModel
                    {
                        Data = ms.ToArray(),
                        Height = image.Height,
                        Width = image.Width,
                        HorizontalResolution = (float)image.MetaData.HorizontalResolution,
                        VerticalResolution = (float)image.MetaData.HorizontalResolution,
                    };
                    sheet.Drawings.AddPicture(Guid.NewGuid().ToString(), imageModel);
                }

                templatePackage.SaveAs(new FileInfo(Guid.NewGuid().ToString() + ".xlsx"));
            }
        }
    }
}
