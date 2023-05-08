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
 *******************************************************************************
 * Jan Källman		Added		10-SEP-2009
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.IO;
using System.Xml;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils;
using SkiaSharp;

namespace OfficeOpenXml
{
    /// <summary>
    /// An image that fills the background of the worksheet.
    /// </summary>
    public class ExcelBackgroundImage : XmlHelper
    {
        ExcelWorksheet _workSheet;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="nsm"></param>
        /// <param name="topNode">The topnode of the worksheet</param>
        /// <param name="workSheet">Worksheet reference</param>
        internal ExcelBackgroundImage(XmlNamespaceManager nsm, XmlNode topNode, ExcelWorksheet workSheet) :
            base(nsm, topNode)
        {
            _workSheet = workSheet;
        }

        private const string BACKGROUNDPIC_PATH = "d:picture/@r:id";
        /// <summary>
        /// The background image of the worksheet. 
        /// The image will be saved internally as a jpg.
        /// </summary>
        public SKImage Image
        {
            get
            {
                var relID = GetXmlNodeString(BACKGROUNDPIC_PATH);
                if (!string.IsNullOrEmpty(relID))
                {
                    var rel = _workSheet.Part.GetRelationship(relID);
                    var imagePart = _workSheet.Part.Package.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));

                    return SKImage.FromEncodedData(imagePart.GetStream());
                }
                return null;
            }
            set
            {
                DeletePrevImage();
                if (value == null)
                {
                    DeleteAllNode(BACKGROUNDPIC_PATH);
                }
                else
                {
                    var img = ImageCompat.GetImageAsByteArray(value);

                    var ii = _workSheet.Workbook._package.AddImage(img);
                    var rel = _workSheet.Part.CreateRelationship(ii.Uri, Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
                    SetXmlNodeString(BACKGROUNDPIC_PATH, rel.Id);
                }
            }
        }
        /// <summary>
        /// Set the picture from an image file. 
        /// The image file will be saved as a blob, so make sure Excel supports the image format.
        /// </summary>
        /// <param name="pictureFile">The image file.</param>
        public void SetFromFile(FileInfo pictureFile)
        {
            DeletePrevImage();

            SKImage img;
            byte[] fileBytes;
            try
            {
                fileBytes = File.ReadAllBytes(pictureFile.FullName);
                img = SKImage.FromEncodedData(pictureFile.FullName);
            }
            catch (Exception ex)
            {
                throw new InvalidDataException("File is not a supported image-file or is corrupt", ex);
            }

            var contentType = ExcelPicture.GetContentType(pictureFile.Extension);
            var imageURI = GetNewUri(_workSheet._package.Package, "/xl/media/" + pictureFile.Name.Substring(0, pictureFile.Name.Length - pictureFile.Extension.Length) + "{0}" + pictureFile.Extension);

            var ii = _workSheet.Workbook._package.AddImage(fileBytes, imageURI, contentType);


            if (_workSheet.Part.Package.PartExists(imageURI) && ii.RefCount==1) //The file exists with another content, overwrite it.
            {
                //Remove the part if it exists
                _workSheet.Part.Package.DeletePart(imageURI);
            }

            var imagePart = _workSheet.Part.Package.CreatePart(imageURI, contentType, CompressionLevel.None);
            //Save the picture to package.

            var strm = imagePart.GetStream(FileMode.Create, FileAccess.Write);
            strm.Write(fileBytes, 0, fileBytes.Length);

            var rel = _workSheet.Part.CreateRelationship(imageURI, Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
            SetXmlNodeString(BACKGROUNDPIC_PATH, rel.Id);
        }
        private void DeletePrevImage()
        {
            var relID = GetXmlNodeString(BACKGROUNDPIC_PATH);
            if (relID != "")
            {
                var img = ImageCompat.GetImageAsByteArray(Image);

                var ii = _workSheet.Workbook._package.GetImageInfo(img);

                //Delete the relation
                _workSheet.Part.DeleteRelationship(relID);

                //Delete the image if there are no other references.
                if (ii != null && ii.RefCount == 1)
                {
                    if (_workSheet.Part.Package.PartExists(ii.Uri))
                    {
                        _workSheet.Part.Package.DeletePart(ii.Uri);
                    }
                }

            }
        }
    }
}
