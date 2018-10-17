using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeOpenXml.Drawing.Custom
{
    public class ImageModel
    {
        public byte[] Data { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public float VerticalResolution { get; set; }
        public float HorizontalResolution { get; set; }

        internal void Save(Stream stream)
        {
            using (var writer = new BinaryWriter(stream))
            {
                writer.Write(this.Data);
            }
        }
    }
}
