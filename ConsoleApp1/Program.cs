using OfficeOpenXml;
using OfficeOpenXml.Drawing.Custom;
using SixLabors.ImageSharp;
using System;
using System.Drawing;
using System.IO;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Run();
            Console.WriteLine("Hello World!");
        }

        private static float MeasureString(string s, Font font)
        {
            using (var g = Graphics.FromHwnd(IntPtr.Zero))
            {
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                return g.MeasureString(s, font, int.MaxValue, StringFormat.GenericTypographic).Width;
            }
        }

        internal static int GetWidthInPixels(ExcelRangeBase cell)
        {
            double columnWidth = cell.Worksheet.Column(cell.Start.Column).Width;
            Font font = new Font(cell.Style.Font.Name, cell.Style.Font.Size, FontStyle.Regular);

            double pxBaseline = Math.Round(MeasureString("1234567890", font) / 10);

            return (int)(columnWidth * pxBaseline);
        }

        internal static int GetHeightInPixels(ExcelRangeBase cell)
        {
            using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                float dpiY = graphics.DpiY;
                return (int)(cell.Worksheet.Row(cell.Start.Row).Height * (1 / 72.0) * dpiY);
            }
        }

        private static void Run()
        {
            using (ExcelPackage templatePackage = new ExcelPackage(new System.IO.FileInfo(@"C:\Temp\template.xlsx")))
            {
                var image = SixLabors.ImageSharp.Image.Load(@"C:\Temp\temp.jpg");
                using (MemoryStream ms = new MemoryStream())
                {
                    image.SaveAsJpeg(ms);
                    var sheet = templatePackage.Workbook.Worksheets[0];

                    var targetCell = sheet.Cells[1, 1];
                    var imageModel = new ImageModel
                    {
                        Data = ms.ToArray(),
                        Height = image.Height,
                        Width = image.Width,
                        HorizontalResolution = (float)image.MetaData.HorizontalResolution,
                        VerticalResolution = (float)image.MetaData.HorizontalResolution,
                    };
                    var picture = sheet.Drawings.AddPicture(Guid.NewGuid().ToString(), imageModel);

                    picture.From.Column = targetCell.Start.Column - 1;
                    picture.From.Row = targetCell.Start.Row - 1;

                    var pixelWidth = targetCell.Worksheet.Column(targetCell.Start.Column).Width * 7;
                    var pixelHeight = targetCell.Worksheet.Row(targetCell.Start.Row).Height * (1 + 1.0 / 3);

                    double multiplier = Math.Min(pixelWidth / (double)image.Width, pixelHeight / (double)image.Height);

                    picture.SetSize((int)(image.Width * multiplier), (int)(image.Height * multiplier));

                }

                templatePackage.SaveAs(new FileInfo(Guid.NewGuid().ToString() + ".xlsx"));
            }
        }
    }
}
