﻿/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied.
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 *
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using System.IO;
using System.Diagnostics;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Compatibility;

namespace OfficeOpenXml.Drawing.Custom
{
    /// <summary>
    /// An image object
    /// </summary>
    public sealed class ExcelPicture : ExcelDrawing
    {
        #region "Constructors"


        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, ImageModel image, Uri hyperlink) :
            base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
        {
            XmlElement picNode = node.OwnerDocument.CreateElement("xdr", "pic", ExcelPackage.schemaSheetDrawings);
            node.InsertAfter(picNode, node.SelectSingleNode("xdr:to", NameSpaceManager));
            _hyperlink = hyperlink;
            picNode.InnerXml = PicStartXml();

            node.InsertAfter(node.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings), picNode);

            var package = drawings.Worksheet._package.Package;
            //Get the picture if it exists or save it if not.
            _image = image;
            string relID = SavePicture(image.Data);

            //Create relationship
            node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relID;
            _height = image.Height;
            _width = image.Width;
            SetPosDefaults(image);
            package.Flush();
        }
        

        internal static string GetContentType(string extension)
        {
            switch (extension.ToLower(CultureInfo.InvariantCulture))
            {
                case ".bmp":
                    return "image/bmp";

                case ".jpg":
                case ".jpeg":
                    return "image/jpeg";

                case ".gif":
                    return "image/gif";

                case ".png":
                    return "image/png";

                case ".cgm":
                    return "image/cgm";

                case ".emf":
                    return "image/x-emf";

                case ".eps":
                    return "image/x-eps";

                case ".pcx":
                    return "image/x-pcx";

                case ".tga":
                    return "image/x-tga";

                case ".tif":
                case ".tiff":
                    return "image/x-tiff";

                case ".wmf":
                    return "image/x-wmf";

                default:
                    return "image/jpeg";
            }
        }

        //internal static ImageFormat GetImageFormat(string contentType)
        //{
        //    switch (contentType.ToLower(CultureInfo.InvariantCulture))
        //    {
        //        case "image/bmp":
        //            return ImageFormat.Bmp;

        //        case "image/jpeg":
        //            return ImageFormat.Jpeg;

        //        case "image/gif":
        //            return ImageFormat.Gif;

        //        case "image/png":
        //            return ImageFormat.Png;

        //        case "image/x-emf":
        //            return ImageFormat.Emf;

        //        case "image/x-tiff":
        //            return ImageFormat.Tiff;

        //        case "image/x-wmf":
        //            return ImageFormat.Wmf;

        //        default:
        //            return ImageFormat.Jpeg;
        //    }
        //}        //Add a new image to the compare collection

        private void AddNewPicture(byte[] img, string relID)
        {
            var newPic = new ExcelDrawings.ImageCompare();
            newPic.image = img;
            newPic.relID = relID;
            //_drawings._pics.Add(newPic);
        }

        #endregion "Constructors"

        private string SavePicture(byte[] img)
        {
            var ii = _drawings._package.AddImage(img);

            ImageHash = ii.Hash;
            if (_drawings._hashes.ContainsKey(ii.Hash))
            {
                var relID = _drawings._hashes[ii.Hash];
                var rel = _drawings.Part.GetRelationship(relID);
                UriPic = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                return relID;
            }
            else
            {
                UriPic = ii.Uri;
                ImageHash = ii.Hash;
            }

            //Set the Image and save it to the package.
            RelPic = _drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(_drawings.UriDrawing, UriPic), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");

            //AddNewPicture(img, picRelation.Id);
            _drawings._hashes.Add(ii.Hash, RelPic.Id);

            return RelPic.Id;
        }
        
        private void SetPosDefaults(ImageModel image)
        {
            EditAs = eEditAs.OneCell;
            SetPixelWidth(image.Width, image.HorizontalResolution);
            SetPixelHeight(image.Height, image.VerticalResolution);
        }

        private string PicStartXml()
        {
            StringBuilder xml = new StringBuilder();

            xml.Append("<xdr:nvPicPr>");

            if (_hyperlink == null)
            {
                xml.AppendFormat("<xdr:cNvPr id=\"{0}\" descr=\"\" />", _id);
            }
            else
            {
                HypRel = _drawings.Part.CreateRelationship(_hyperlink, Packaging.TargetMode.External, ExcelPackage.schemaHyperlink);
                xml.AppendFormat("<xdr:cNvPr id=\"{0}\" descr=\"\">", _id);
                if (HypRel != null)
                {
                    if (_hyperlink is ExcelHyperLink)
                    {
                        xml.AppendFormat("<a:hlinkClick xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{0}\" tooltip=\"{1}\"/>",
                          HypRel.Id, ((ExcelHyperLink)_hyperlink).ToolTip);
                    }
                    else
                    {
                        xml.AppendFormat("<a:hlinkClick xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{0}\" />",
                          HypRel.Id);
                    }
                }
                xml.Append("</xdr:cNvPr>");
            }

            xml.Append("<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"\" cstate=\"print\" /><a:stretch><a:fillRect /> </a:stretch> </xdr:blipFill> <xdr:spPr> <a:xfrm> <a:off x=\"0\" y=\"0\" />  <a:ext cx=\"0\" cy=\"0\" /> </a:xfrm> <a:prstGeom prst=\"rect\"> <a:avLst /> </a:prstGeom> </xdr:spPr>");

            return xml.ToString();
        }

        internal string ImageHash { get; set; }
        private ImageModel _image = null;

        /// <summary>
        /// The Image
        /// </summary>
        public ImageModel ImageModel
        {
            get
            {
                return _image;
            }
            set
            {
                if (value != null)
                {
                    _image = value;
                    try
                    {
                        string relID = SavePicture(value.Data);

                        //Create relationship
                        TopNode.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relID;
                        //_image.Save(Part.GetStream(FileMode.Create, FileAccess.Write), _imageFormat);   //Always JPEG here at this point.
                    }
                    catch (Exception ex)
                    {
                        throw (new Exception("Can't save image - " + ex.Message, ex));
                    }
                }
            }
        }

        //private ImageFormat _imageFormat = ImageFormat.Jpeg;

        /// <summary>
        /// Image format
        /// If the picture is created from an Image this type is always Jpeg
        /// </summary>
        //public ImageFormat ImageFormat
        //{
        //    get
        //    {
        //        return _imageFormat;
        //    }
        //    internal set
        //    {
        //        _imageFormat = value;
        //    }
        //}

        //internal string ContentType
        //{
        //    get;
        //    set;
        //}

        /// <summary>
        /// Set the size of the image in percent from the orginal size
        /// Note that resizing columns / rows after using this function will effect the size of the picture
        /// </summary>
        /// <param name="Percent">Percent</param>
        public override void SetSize(int Percent)
        {
            if (ImageModel == null)
            {
                base.SetSize(Percent);
            }
            else
            {
                _width = ImageModel.Width;
                _height = ImageModel.Height;

                _width = (int)(_width * ((decimal)Percent / 100));
                _height = (int)(_height * ((decimal)Percent / 100));

                SetPixelWidth(_width, ImageModel.HorizontalResolution);
                SetPixelHeight(_height, ImageModel.VerticalResolution);
            }
        }

        internal Uri UriPic { get; set; }
        internal Packaging.ZipPackageRelationship RelPic { get; set; }
        internal Packaging.ZipPackageRelationship HypRel { get; set; }
        internal Packaging.ZipPackagePart Part;

        internal new string Id
        {
            get { return Name; }
        }

        private ExcelDrawingFill _fill = null;

        /// <summary>
        /// Fill
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(NameSpaceManager, TopNode, "xdr:pic/xdr:spPr");
                }
                return _fill;
            }
        }

        private ExcelDrawingBorder _border = null;

        /// <summary>
        /// Border
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(NameSpaceManager, TopNode, "xdr:pic/xdr:spPr/a:ln");
                }
                return _border;
            }
        }

        private Uri _hyperlink = null;

        /// <summary>
        /// Hyperlink
        /// </summary>
        public Uri Hyperlink
        {
            get
            {
                return _hyperlink;
            }
        }

        internal override void DeleteMe()
        {
            _drawings._package.RemoveImage(ImageHash);
            base.DeleteMe();
        }

        public override void Dispose()
        {
            base.Dispose();
            _hyperlink = null;
            _image = null;
        }
    }
}