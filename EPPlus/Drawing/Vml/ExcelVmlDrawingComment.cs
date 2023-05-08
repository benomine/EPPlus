/*******************************************************************************
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
 * Jan Källman		Initial Release		        2010-06-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Globalization;
using System.Xml;
using SkiaSharp;

namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// Drawing object used for comments
    /// </summary>
    public class ExcelVmlDrawingComment : ExcelVmlDrawingBase, IRangeID
    {
        internal ExcelVmlDrawingComment(XmlNode topNode, ExcelRangeBase range, XmlNamespaceManager ns) :
            base(topNode, ns)
        {
            Range = range;
            SchemaNodeOrder = new string[] { "fill", "stroke", "shadow", "path", "textbox", "ClientData", "MoveWithCells", "SizeWithCells", "Anchor", "Locked", "AutoFill", "LockText", "TextHAlign", "TextVAlign", "Row", "Column", "Visible" };
        }
        internal ExcelRangeBase Range { get; set; }

        /// <summary>
        /// Address in the worksheet
        /// </summary>
        public string Address => Range.Address;

        const string VERTICAL_ALIGNMENT_PATH = "x:ClientData/x:TextVAlign";
        /// <summary>
        /// Vertical alignment for text
        /// </summary>
        public eTextAlignVerticalVml VerticalAlignment
        {
            get
            {
                return GetXmlNodeString(VERTICAL_ALIGNMENT_PATH) switch
                {
                    "Center" => eTextAlignVerticalVml.Center,
                    "Bottom" => eTextAlignVerticalVml.Bottom,
                    _ => eTextAlignVerticalVml.Top
                };
            }
            set
            {
                switch (value)
                {
                    case eTextAlignVerticalVml.Center:
                        SetXmlNodeString(VERTICAL_ALIGNMENT_PATH, "Center");
                        break;
                    case eTextAlignVerticalVml.Bottom:
                        SetXmlNodeString(VERTICAL_ALIGNMENT_PATH, "Bottom");
                        break;
                    default:
                        DeleteNode(VERTICAL_ALIGNMENT_PATH);
                        break;
                }
            }
        }
        const string HORIZONTAL_ALIGNMENT_PATH = "x:ClientData/x:TextHAlign";
        /// <summary>
        /// Horizontal alignment for text
        /// </summary>
        public eTextAlignHorizontalVml HorizontalAlignment
        {
            get
            {
                return GetXmlNodeString(HORIZONTAL_ALIGNMENT_PATH) switch
                {
                    "Center" => eTextAlignHorizontalVml.Center,
                    "Right" => eTextAlignHorizontalVml.Right,
                    _ => eTextAlignHorizontalVml.Left
                };
            }
            set
            {
                switch (value)
                {
                    case eTextAlignHorizontalVml.Center:
                        SetXmlNodeString(HORIZONTAL_ALIGNMENT_PATH, "Center");
                        break;
                    case eTextAlignHorizontalVml.Right:
                        SetXmlNodeString(HORIZONTAL_ALIGNMENT_PATH, "Right");
                        break;
                    default:
                        DeleteNode(HORIZONTAL_ALIGNMENT_PATH);
                        break;
                }
            }
        }
        const string VISIBLE_PATH = "x:ClientData/x:Visible";
        /// <summary>
        /// If the drawing object is visible.
        /// </summary>
        public bool Visible
        {
            get => TopNode.SelectSingleNode(VISIBLE_PATH, NameSpaceManager)!=null;
            set
            {
                if (value)
                {
                    CreateNode(VISIBLE_PATH);
                    Style = SetStyle(Style, "visibility", "visible");
                }
                else
                {
                    DeleteNode(VISIBLE_PATH);
                    Style = SetStyle(Style, "visibility", "hidden");
                }
            }
        }

        const string BACKGROUNDCOLOR_PATH = "@fillcolor";
        const string BACKGROUNDCOLOR2_PATH = "v:fill/@color2";
        /// <summary>
        /// Background color
        /// </summary>
        public SKColor BackgroundColor
        {
            get
            {
                var col = GetXmlNodeString(BACKGROUNDCOLOR_PATH);
                if (col == "")
                {
                    return SKColor.FromHsl(0xff, 0xff, 0xe1);
                }

                if (col.StartsWith("#")) col=col.Substring(1, col.Length-1);
                if (int.TryParse(col, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var res))
                {
                    return SKColor.Parse(col);
                }

                return SKColors.Empty;
            }
            set
            {
                var color = "#" + value.ToString().Substring(2, 6);
                SetXmlNodeString(BACKGROUNDCOLOR_PATH, color);
            }
        }
        const string LINESTYLE_PATH = "v:stroke/@dashstyle";
        const string ENDCAP_PATH = "v:stroke/@endcap";
        /// <summary>
        /// Linestyle for border
        /// </summary>
        public eLineStyleVml LineStyle
        {
            get
            {
                var v = GetXmlNodeString(LINESTYLE_PATH);
                switch (v)
                {
                    case "":
                        return eLineStyleVml.Solid;
                    case "1 1":
                        v = GetXmlNodeString(ENDCAP_PATH);
                        return (eLineStyleVml)Enum.Parse(typeof(eLineStyleVml), v, true);
                    default:
                        return (eLineStyleVml)Enum.Parse(typeof(eLineStyleVml), v, true);
                }
            }
            set
            {
                if (value == eLineStyleVml.Round || value == eLineStyleVml.Square)
                {
                    SetXmlNodeString(LINESTYLE_PATH, "1 1");
                    if (value == eLineStyleVml.Round)
                    {
                        SetXmlNodeString(ENDCAP_PATH, "round");
                    }
                    else
                    {
                        DeleteNode(ENDCAP_PATH);
                    }
                }
                else
                {
                    var v = value.ToString();
                    v = v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1, v.Length - 1);
                    SetXmlNodeString(LINESTYLE_PATH, v);
                    DeleteNode(ENDCAP_PATH);
                }
            }
        }
        const string LINECOLOR_PATH = "@strokecolor";
        /// <summary>
        /// Line color 
        /// </summary>
        public SKColor LineColor
        {
            get
            {
                var col = GetXmlNodeString(LINECOLOR_PATH);
                if (col == "")
                {
                    return SKColors.Black;
                }

                if (col.StartsWith("#")) col = col.Substring(1, col.Length - 1);
                if (int.TryParse(col, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var res))
                {
                    return SKColor.Parse(col);
                }

                return SKColors.Empty;
            }
            set
            {
                var color = "#" + value.ToString().Substring(2, 6);
                SetXmlNodeString(LINECOLOR_PATH, color);
            }
        }
        const string LINEWIDTH_PATH = "@strokeweight";
        /// <summary>
        /// Width of the border
        /// </summary>
        public Single LineWidth
        {
            get
            {
                var wt = GetXmlNodeString(LINEWIDTH_PATH);
                if (wt == "") return (Single).75;
                if (wt.EndsWith("pt")) wt=wt.Substring(0, wt.Length-2);

                if (float.TryParse(wt, NumberStyles.Any, CultureInfo.InvariantCulture, out var ret))
                {
                    return ret;
                }

                return 0;
            }
            set => SetXmlNodeString(LINEWIDTH_PATH, value.ToString(CultureInfo.InvariantCulture) + "pt");
        }

        const string TEXTBOX_STYLE_PATH = "v:textbox/@style";
        /// <summary>
        /// Autofits the drawingobject 
        /// </summary>
        public bool AutoFit
        {
            get
            {
                string value;
                GetStyle(GetXmlNodeString(TEXTBOX_STYLE_PATH), "mso-fit-shape-to-text", out value);
                return value=="t";
            }
            set => SetXmlNodeString(TEXTBOX_STYLE_PATH, SetStyle(GetXmlNodeString(TEXTBOX_STYLE_PATH), "mso-fit-shape-to-text", value ? "t" : ""));
        }
        const string LOCKED_PATH = "x:ClientData/x:Locked";

        public bool Locked
        {
            get => GetXmlNodeBool(LOCKED_PATH, false);
            set => SetXmlNodeBool(LOCKED_PATH, value, false);
        }
        const string LOCK_TEXT_PATH = "x:ClientData/x:LockText";

        public bool LockText
        {
            get => GetXmlNodeBool(LOCK_TEXT_PATH, false);
            set => SetXmlNodeBool(LOCK_TEXT_PATH, value, false);
        }
        ExcelVmlDrawingPosition _from = null;

        public ExcelVmlDrawingPosition From =>
            _from ??= new ExcelVmlDrawingPosition(NameSpaceManager,
                TopNode.SelectSingleNode("x:ClientData", NameSpaceManager), 0);

        ExcelVmlDrawingPosition _to = null;

        public ExcelVmlDrawingPosition To =>
            _to ??= new ExcelVmlDrawingPosition(NameSpaceManager,
                TopNode.SelectSingleNode("x:ClientData", NameSpaceManager), 4);

        const string ROW_PATH = "x:ClientData/x:Row";

        internal int Row
        {
            get => GetXmlNodeInt(ROW_PATH);
            set => SetXmlNodeString(ROW_PATH, value.ToString(CultureInfo.InvariantCulture));
        }
        const string COLUMN_PATH = "x:ClientData/x:Column";

        internal int Column
        {
            get => GetXmlNodeInt(COLUMN_PATH);
            set => SetXmlNodeString(COLUMN_PATH, value.ToString(CultureInfo.InvariantCulture));
        }
        const string STYLE_PATH = "@style";
        public string Style
        {
            get => GetXmlNodeString(STYLE_PATH);
            set => SetXmlNodeString(STYLE_PATH, value);
        }
        #region IRangeID Members

        ulong IRangeID.RangeID
        {
            get => ExcelCellBase.GetCellID(Range.Worksheet.SheetID, Range.Start.Row, Range.Start.Column);
            set
            {

            }
        }

        #endregion
    }
}
