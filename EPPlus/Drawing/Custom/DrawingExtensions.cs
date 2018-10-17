using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Drawing.Custom;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace OfficeOpenXml.Drawing
{
    public static class DrawingExtensions
    {
        public static OfficeOpenXml.Drawing.Custom.ExcelPicture AddPicture(this OfficeOpenXml.Drawing.ExcelDrawings excelDrawings, string name, Image image)
        {
            return excelDrawings.AddPicture(name, image, null);
        }
        public static OfficeOpenXml.Drawing.Custom.ExcelPicture AddPicture(this OfficeOpenXml.Drawing.ExcelDrawings excelDrawings, string name, Image image,Uri link )
        {
#if (Core)
            byte[] img = ImageCompat.GetImageAsByteArray(image);
#else
                        ImageConverter ic = new ImageConverter();
                        byte[] img = (byte[])ic.ConvertTo(image, typeof(byte[]));
#endif

            return excelDrawings.AddPicture(name, new OfficeOpenXml.Drawing.Custom.ImageModel
            {
                Data = img,
                Height = image.Height,
                Width = image.Width,
                HorizontalResolution =
                image.HorizontalResolution,
                VerticalResolution = image.HorizontalResolution
            },link);
        }
    }
}
