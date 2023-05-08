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
 * Jan Källman		                Initial Release		        2009-12-22
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Globalization;
using System.Xml;
using SkiaSharp;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Fill properties for drawing objects
    /// </summary>
    public sealed class ExcelDrawingFill : XmlHelper
    {
        string _fillPath;
        XmlNode _fillNode;
        internal ExcelDrawingFill(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string fillPath) :
            base(nameSpaceManager, topNode)
        {
            _fillPath = fillPath;
            _fillNode = topNode.SelectSingleNode(_fillPath, NameSpaceManager);
            SchemaNodeOrder = new string[] { "tickLblPos", "spPr", "txPr", "dLblPos", "crossAx", "printSettings", "showVal", "prstGeom", "noFill", "solidFill", "blipFill", "gradFill", "noFill", "pattFill", "ln", "prstDash" };
            if (_fillNode != null)
            {
                _fillTypeNode = topNode.SelectSingleNode("solidFill");
                if (_fillTypeNode == null) _fillTypeNode = topNode.SelectSingleNode("noFill");
                if (_fillTypeNode == null) _fillTypeNode = topNode.SelectSingleNode("blipFill");
                if (_fillTypeNode == null) _fillTypeNode = topNode.SelectSingleNode("gradFill");
                if (_fillTypeNode == null) _fillTypeNode = topNode.SelectSingleNode("pattFill");
            }
        }
        eFillStyle _style;
        XmlNode _fillTypeNode = null;
        /// <summary>
        /// Fill style
        /// </summary>
        public eFillStyle Style
        {
            get
            {
                if (_fillTypeNode == null)
                {
                    return eFillStyle.SolidFill;
                }

                _style=GetStyleEnum(_fillTypeNode.Name);
                return _style;
            }
            set
            {
                if (value == eFillStyle.NoFill || value == eFillStyle.SolidFill)
                {
                    _style = value;
                    CreateFillTopNode(value);
                }
                else
                {
                    throw new NotImplementedException("Fillstyle not implemented");
                }
            }
        }

        private void CreateFillTopNode(eFillStyle value)
        {
            if (_fillTypeNode != null)
            {
                TopNode.RemoveChild(_fillTypeNode);
            }
            CreateNode(_fillPath + "/a:" + GetStyleText(value), false);
            _fillNode=TopNode.SelectSingleNode(_fillPath + "/a:" + GetStyleText(value), NameSpaceManager);
        }

        private eFillStyle GetStyleEnum(string name)
        {
            switch (name)
            {
                case "noFill":
                    return eFillStyle.NoFill;
                case "blipFill":
                    return eFillStyle.BlipFill;
                case "gradFill":
                    return eFillStyle.GradientFill;
                case "grpFill":
                    return eFillStyle.GroupFill;
                case "pattFill":
                    return eFillStyle.PatternFill;
                default:
                    return eFillStyle.SolidFill;
            }
        }

        private string GetStyleText(eFillStyle style)
        {
            switch (style)
            {
                case eFillStyle.BlipFill:
                    return "blipFill";
                case eFillStyle.GradientFill:
                    return "gradFill";
                case eFillStyle.GroupFill:
                    return "grpFill";
                case eFillStyle.NoFill:
                    return "noFill";
                case eFillStyle.PatternFill:
                    return "pattFill";
                default:
                    return "solidFill";
            }
        }

        const string ColorPath = "/a:solidFill/a:srgbClr/@val";
        /// <summary>
        /// Fill color for solid fills
        /// </summary>
        public SKColor Color
        {
            get
            {
                var col = GetXmlNodeString(_fillPath + ColorPath);
                if (col == "")
                {
                    return new SKColor(79, 129, 189);
                }

                return new SKColor((uint)int.Parse(col, NumberStyles.AllowHexSpecifier));
            }
            set
            {
                if (_fillTypeNode == null)
                {
                    _style = eFillStyle.SolidFill;
                }
                else if (_style != eFillStyle.SolidFill)
                {
                    throw new Exception("FillStyle must be set to SolidFill");
                }
                CreateNode(_fillPath, false);
                SetXmlNodeString(_fillPath + ColorPath, value.ToString().Substring(3));
            }
        }
        const string alphaPath = "/a:solidFill/a:srgbClr/a:alpha/@val";
        /// <summary>
        /// Transparancy in percent
        /// </summary>
        public int Transparancy
        {
            get => 100 - GetXmlNodeInt(_fillPath + alphaPath) / 1000;
            set
            {
                if (_fillTypeNode == null)
                {
                    _style = eFillStyle.SolidFill;
                    Color = new SKColor(79, 129, 189);
                }
                else if (_style != eFillStyle.SolidFill)
                {
                    throw new Exception("FillStyle must be set to SolidFill");
                }
                SetXmlNodeString(_fillPath + alphaPath, ((100 - value) * 1000).ToString());
            }
        }
    }
}
