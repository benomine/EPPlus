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
 *******************************************************************************
 * Jan Källman		Added		2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/

using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Bar chart
    /// </summary>
    public sealed class ExcelBarChart : ExcelChart
    {
        #region "Constructors"

        internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart,
            ExcelPivotTable pivotTableSource) :
            base(drawings, node, type, topChart, pivotTableSource)
        {
            SetChartNodeText("");

            SetTypeProperties(drawings, type);
        }

        internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part,
            XmlDocument chartXml, XmlNode chartNode) :
            base(drawings, node, uriChart, part, chartXml, chartNode)
        {
            SetChartNodeText(chartNode.Name);
        }

        internal ExcelBarChart(ExcelChart topChart, XmlNode chartNode) :
            base(topChart, chartNode)
        {
            SetChartNodeText(chartNode.Name);
        }

        #endregion

        #region "Private functions"

        private void SetChartNodeText(string chartNodeText)
        {
            if (string.IsNullOrEmpty(chartNodeText))
            {
                chartNodeText = GetChartNodeText();
            }
        }

        private void SetTypeProperties(ExcelDrawings drawings, eChartType type)
        {
            if (type is eChartType.BarClustered or eChartType.BarStacked or eChartType.BarStacked100
                or eChartType.BarClustered3D or eChartType.BarStacked3D or eChartType.BarStacked1003D
                or eChartType.ConeBarClustered or eChartType.ConeBarStacked or eChartType.ConeBarStacked100
                or eChartType.CylinderBarClustered or eChartType.CylinderBarStacked or eChartType.CylinderBarStacked100
                or eChartType.PyramidBarClustered or eChartType.PyramidBarStacked or eChartType.PyramidBarStacked100)
            {
                Direction = eDirection.Bar;
            }
            else if (
                type is eChartType.ColumnClustered or eChartType.ColumnStacked or eChartType.ColumnStacked100
                or eChartType.Column3D or eChartType.ColumnClustered3D or eChartType.ColumnStacked3D
                or eChartType.ColumnStacked1003D or eChartType.ConeCol or eChartType.ConeColClustered
                or eChartType.ConeColStacked or eChartType.ConeColStacked100 or eChartType.CylinderCol
                or eChartType.CylinderColClustered or eChartType.CylinderColStacked or eChartType.CylinderColStacked100
                or eChartType.PyramidCol or eChartType.PyramidColClustered or eChartType.PyramidColStacked
                or eChartType.PyramidColStacked100)
            {
                Direction = eDirection.Column;
            }

            if (
                type is eChartType.Column3D or eChartType.ColumnClustered3D or eChartType.ColumnStacked3D
                or eChartType.ColumnStacked1003D or eChartType.BarClustered3D or eChartType.BarStacked3D
                or eChartType.BarStacked1003D)
            {
                Shape = eShape.Box;
            }
            else if (
                type is eChartType.CylinderBarClustered or eChartType.CylinderBarStacked
                or eChartType.CylinderBarStacked100 or eChartType.CylinderCol or eChartType.CylinderColClustered
                or eChartType.CylinderColStacked or eChartType.CylinderColStacked100)
            {
                Shape = eShape.Cylinder;
            }
            else if (
                type is eChartType.ConeBarClustered or eChartType.ConeBarStacked or eChartType.ConeBarStacked100
                or eChartType.ConeCol or eChartType.ConeColClustered or eChartType.ConeColStacked
                or eChartType.ConeColStacked100)
            {
                Shape = eShape.Cone;
            }
            else if (
                type is eChartType.PyramidBarClustered or eChartType.PyramidBarStacked
                or eChartType.PyramidBarStacked100 or eChartType.PyramidCol or eChartType.PyramidColClustered
                or eChartType.PyramidColStacked or eChartType.PyramidColStacked100)
            {
                Shape = eShape.Pyramid;
            }
        }

        #endregion

        #region "Properties"

        string _directionPath = "c:barDir/@val";

        /// <summary>
        /// Direction, Bar or columns
        /// </summary>
        public eDirection Direction
        {
            get => GetDirectionEnum(_chartXmlHelper.GetXmlNodeString(_directionPath));
            internal set => _chartXmlHelper.SetXmlNodeString(_directionPath, GetDirectionText(value));
        }

        string _shapePath = "c:shape/@val";

        /// <summary>
        /// The shape of the bar/columns
        /// </summary>
        public eShape Shape
        {
            get => GetShapeEnum(_chartXmlHelper.GetXmlNodeString(_shapePath));
            internal set => _chartXmlHelper.SetXmlNodeString(_shapePath, GetShapeText(value));
        }

        ExcelChartDataLabel _DataLabel = null;

        /// <summary>
        /// Access to datalabel properties
        /// </summary>
        public ExcelChartDataLabel DataLabel => _DataLabel ??= new ExcelChartDataLabel(NameSpaceManager, ChartNode);

        string _gapWidthPath = "c:gapWidth/@val";

        /// <summary>
        /// The size of the gap between two adjacent bars/columns
        /// </summary>
        public int GapWidth
        {
            get => _chartXmlHelper.GetXmlNodeInt(_gapWidthPath);
            set => _chartXmlHelper.SetXmlNodeString(_gapWidthPath, value.ToString(CultureInfo.InvariantCulture));
        }

        #endregion

        #region "Direction Enum Traslation"

        private string GetDirectionText(eDirection direction)
        {
            return direction switch
            {
                eDirection.Bar => "bar",
                _ => "col"
            };
        }

        private eDirection GetDirectionEnum(string direction)
        {
            return direction switch
            {
                "bar" => eDirection.Bar,
                _ => eDirection.Column
            };
        }

        #endregion

        #region "shape Enum Translation"

        private string GetShapeText(eShape shape)
        {
            return shape switch
            {
                eShape.Box => "box",
                eShape.Cone => "cone",
                eShape.ConeToMax => "coneToMax",
                eShape.Cylinder => "cylinder",
                eShape.Pyramid => "pyramid",
                eShape.PyramidToMax => "pyramidToMax",
                _ => "box"
            };
        }

        private eShape GetShapeEnum(string text)
        {
            return text switch
            {
                "box" => eShape.Box,
                "cone" => eShape.Cone,
                "coneToMax" => eShape.ConeToMax,
                "cylinder" => eShape.Cylinder,
                "pyramid" => eShape.Pyramid,
                "pyramidToMax" => eShape.PyramidToMax,
                _ => eShape.Box
            };
        }

        #endregion

        internal override eChartType GetChartType(string name)
        {
            if (name == "barChart")
            {
                if (Direction == eDirection.Bar)
                {
                    return Grouping switch
                    {
                        eGrouping.Stacked => eChartType.BarStacked,
                        eGrouping.PercentStacked => eChartType.BarStacked100,
                        _ => eChartType.BarClustered
                    };
                }

                return Grouping switch
                {
                    eGrouping.Stacked => eChartType.ColumnStacked,
                    eGrouping.PercentStacked => eChartType.ColumnStacked100,
                    _ => eChartType.ColumnClustered
                };
            }

            if (name != "bar3DChart") return base.GetChartType(name);

            #region "Bar shape"

            switch (Shape)
            {
                case eShape.Box when Direction == eDirection.Bar:
                {
                    return Grouping switch
                    {
                        eGrouping.Stacked => eChartType.BarStacked3D,
                        eGrouping.PercentStacked => eChartType.BarStacked1003D,
                        _ => eChartType.BarClustered3D
                    };
                }
                case eShape.Box when Grouping == eGrouping.Stacked:
                    return eChartType.ColumnStacked3D;
                case eShape.Box when Grouping == eGrouping.PercentStacked:
                    return eChartType.ColumnStacked1003D;
                case eShape.Box:
                    return eChartType.ColumnClustered3D;
                case eShape.Cone:
                case eShape.ConeToMax:
                {
                    if (Direction == eDirection.Bar)
                    {
                        return Grouping switch
                        {
                            eGrouping.Stacked => eChartType.ConeBarStacked,
                            eGrouping.PercentStacked => eChartType.ConeBarStacked100,
                            eGrouping.Clustered => eChartType.ConeBarClustered
                        };
                    }

                    return Grouping switch
                    {
                        eGrouping.Stacked => eChartType.ConeColStacked,
                        eGrouping.PercentStacked => eChartType.ConeColStacked100,
                        eGrouping.Clustered => eChartType.ConeColClustered,
                        _ => eChartType.ConeCol
                    };
                }
            }

            #endregion

            #region "Cylinder shape"

            if (Shape == eShape.Cylinder)
            {
                if (Direction == eDirection.Bar)
                {
                    return Grouping switch
                    {
                        eGrouping.Stacked => eChartType.CylinderBarStacked,
                        eGrouping.PercentStacked => eChartType.CylinderBarStacked100,
                        eGrouping.Clustered => eChartType.CylinderBarClustered
                    };
                }

                return Grouping switch
                {
                    eGrouping.Stacked => eChartType.CylinderColStacked,
                    eGrouping.PercentStacked => eChartType.CylinderColStacked100,
                    eGrouping.Clustered => eChartType.CylinderColClustered,
                    _ => eChartType.CylinderCol
                };
            }

            #endregion

            #region "Pyramide shape"

            if (Shape is eShape.Pyramid or eShape.PyramidToMax)
            {
                if (Direction == eDirection.Bar)
                {
                    return Grouping switch
                    {
                        eGrouping.Stacked => eChartType.PyramidBarStacked,
                        eGrouping.PercentStacked => eChartType.PyramidBarStacked100,
                        eGrouping.Clustered => eChartType.PyramidBarClustered
                    };
                }

                return Grouping switch
                {
                    eGrouping.Stacked => eChartType.PyramidColStacked,
                    eGrouping.PercentStacked => eChartType.PyramidColStacked100,
                    eGrouping.Clustered => eChartType.PyramidColClustered,
                    _ => eChartType.PyramidCol
                };
            }

            #endregion

            return base.GetChartType(name);
        }
    }
}