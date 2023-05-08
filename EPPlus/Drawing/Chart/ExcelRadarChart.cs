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
 * Jan Källman		Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Provides access to line chart specific properties
    /// </summary>
    public class ExcelRadarChart : ExcelChart
    {
        #region "Constructors"
        internal ExcelRadarChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
            base(drawings, node, uriChart, part, chartXml, chartNode)
        {
            SetTypeProperties();
        }

        internal ExcelRadarChart(ExcelChart topChart, XmlNode chartNode) :
            base(topChart, chartNode)
        {
            SetTypeProperties();
        }
        internal ExcelRadarChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
            base(drawings, node, type, topChart, PivotTableSource)
        {
            SetTypeProperties();
        }
        #endregion
        private void SetTypeProperties()
        {
            RadarStyle = ChartType switch
            {
                eChartType.RadarFilled => eRadarStyle.Filled,
                eChartType.RadarMarkers => eRadarStyle.Marker,
                _ => eRadarStyle.Standard
            };
        }

        string STYLE_PATH = "c:radarStyle/@val";
        /// <summary>
        /// The type of radarchart
        /// </summary>
        public eRadarStyle RadarStyle
        {
            get
            {
                var v = _chartXmlHelper.GetXmlNodeString(STYLE_PATH);
                if (string.IsNullOrEmpty(v))
                {
                    return eRadarStyle.Standard;
                }

                return (eRadarStyle)Enum.Parse(typeof(eRadarStyle), v, true);
            }
            set => _chartXmlHelper.SetXmlNodeString(STYLE_PATH, value.ToString().ToLower(CultureInfo.InvariantCulture));
        }

        ExcelChartDataLabel _DataLabel = null;
        /// <summary>
        /// Access to datalabel properties
        /// </summary>
        public ExcelChartDataLabel DataLabel => _DataLabel ??= new ExcelChartDataLabel(NameSpaceManager, ChartNode);

        internal override eChartType GetChartType(string name)
        {
            return RadarStyle switch
            {
                eRadarStyle.Filled => eChartType.RadarFilled,
                eRadarStyle.Marker => eChartType.RadarMarkers,
                _ => eChartType.Radar
            };
        }
    }
}
