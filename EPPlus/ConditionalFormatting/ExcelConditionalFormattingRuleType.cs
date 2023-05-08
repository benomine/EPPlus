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
 * Author          Change						                  Date
 * ******************************************************************************
 * Eyal Seagull    Conditional Formatting Adaption    2012-04-03
 *******************************************************************************/
using System;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// Functions related to the ExcelConditionalFormattingRule
    /// </summary>
    internal static class ExcelConditionalFormattingRuleType
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="attribute"></param>
        /// <param name="topNode"></param>
        /// <param name="nameSpaceManager"></param>
        /// <returns></returns>
        internal static eExcelConditionalFormattingRuleType GetTypeByAttrbiute(
          string attribute,
          XmlNode topNode,
          XmlNamespaceManager nameSpaceManager)
        {
            switch (attribute)
            {
                case ExcelConditionalFormattingConstants.RuleType.AboveAverage:
                    return GetAboveAverageType(
                      topNode,
                      nameSpaceManager);

                case ExcelConditionalFormattingConstants.RuleType.Top10:
                    return GetTop10Type(
                      topNode,
                      nameSpaceManager);

                case ExcelConditionalFormattingConstants.RuleType.TimePeriod:
                    return GetTimePeriodType(
                      topNode,
                      nameSpaceManager);
                case ExcelConditionalFormattingConstants.RuleType.CellIs:
                    return GetCellIs((XmlElement)topNode);
                case ExcelConditionalFormattingConstants.RuleType.BeginsWith:
                    return eExcelConditionalFormattingRuleType.BeginsWith;

                case ExcelConditionalFormattingConstants.RuleType.Between:
                    return eExcelConditionalFormattingRuleType.Between;

                case ExcelConditionalFormattingConstants.RuleType.ContainsBlanks:
                    return eExcelConditionalFormattingRuleType.ContainsBlanks;

                case ExcelConditionalFormattingConstants.RuleType.ContainsErrors:
                    return eExcelConditionalFormattingRuleType.ContainsErrors;

                case ExcelConditionalFormattingConstants.RuleType.ContainsText:
                    return eExcelConditionalFormattingRuleType.ContainsText;

                case ExcelConditionalFormattingConstants.RuleType.DuplicateValues:
                    return eExcelConditionalFormattingRuleType.DuplicateValues;

                case ExcelConditionalFormattingConstants.RuleType.EndsWith:
                    return eExcelConditionalFormattingRuleType.EndsWith;

                case ExcelConditionalFormattingConstants.RuleType.Equal:
                    return eExcelConditionalFormattingRuleType.Equal;

                case ExcelConditionalFormattingConstants.RuleType.Expression:
                    return eExcelConditionalFormattingRuleType.Expression;

                case ExcelConditionalFormattingConstants.RuleType.GreaterThan:
                    return eExcelConditionalFormattingRuleType.GreaterThan;

                case ExcelConditionalFormattingConstants.RuleType.GreaterThanOrEqual:
                    return eExcelConditionalFormattingRuleType.GreaterThanOrEqual;

                case ExcelConditionalFormattingConstants.RuleType.LessThan:
                    return eExcelConditionalFormattingRuleType.LessThan;

                case ExcelConditionalFormattingConstants.RuleType.LessThanOrEqual:
                    return eExcelConditionalFormattingRuleType.LessThanOrEqual;

                case ExcelConditionalFormattingConstants.RuleType.NotBetween:
                    return eExcelConditionalFormattingRuleType.NotBetween;

                case ExcelConditionalFormattingConstants.RuleType.NotContainsBlanks:
                    return eExcelConditionalFormattingRuleType.NotContainsBlanks;

                case ExcelConditionalFormattingConstants.RuleType.NotContainsErrors:
                    return eExcelConditionalFormattingRuleType.NotContainsErrors;

                case ExcelConditionalFormattingConstants.RuleType.NotContainsText:
                    return eExcelConditionalFormattingRuleType.NotContainsText;

                case ExcelConditionalFormattingConstants.RuleType.NotEqual:
                    return eExcelConditionalFormattingRuleType.NotEqual;

                case ExcelConditionalFormattingConstants.RuleType.UniqueValues:
                    return eExcelConditionalFormattingRuleType.UniqueValues;

                case ExcelConditionalFormattingConstants.RuleType.ColorScale:
                    return GetColorScaleType(
                      topNode,
                      nameSpaceManager);
                case ExcelConditionalFormattingConstants.RuleType.IconSet:
                    return GetIconSetType(topNode, nameSpaceManager);
                case ExcelConditionalFormattingConstants.RuleType.DataBar:
                    return eExcelConditionalFormattingRuleType.DataBar;
            }

            throw new Exception(
              ExcelConditionalFormattingConstants.Errors.UnexpectedRuleTypeAttribute);
        }

        private static eExcelConditionalFormattingRuleType GetCellIs(XmlElement node)
        {
            switch (node.GetAttribute("operator"))
            {
                case ExcelConditionalFormattingConstants.Operators.BeginsWith:
                    return eExcelConditionalFormattingRuleType.BeginsWith;
                case ExcelConditionalFormattingConstants.Operators.Between:
                    return eExcelConditionalFormattingRuleType.Between;

                case ExcelConditionalFormattingConstants.Operators.ContainsText:
                    return eExcelConditionalFormattingRuleType.ContainsText;

                case ExcelConditionalFormattingConstants.Operators.EndsWith:
                    return eExcelConditionalFormattingRuleType.EndsWith;

                case ExcelConditionalFormattingConstants.Operators.Equal:
                    return eExcelConditionalFormattingRuleType.Equal;

                case ExcelConditionalFormattingConstants.Operators.GreaterThan:
                    return eExcelConditionalFormattingRuleType.GreaterThan;

                case ExcelConditionalFormattingConstants.Operators.GreaterThanOrEqual:
                    return eExcelConditionalFormattingRuleType.GreaterThanOrEqual;

                case ExcelConditionalFormattingConstants.Operators.LessThan:
                    return eExcelConditionalFormattingRuleType.LessThan;

                case ExcelConditionalFormattingConstants.Operators.LessThanOrEqual:
                    return eExcelConditionalFormattingRuleType.LessThanOrEqual;

                case ExcelConditionalFormattingConstants.Operators.NotBetween:
                    return eExcelConditionalFormattingRuleType.NotBetween;

                case ExcelConditionalFormattingConstants.Operators.NotContains:
                    return eExcelConditionalFormattingRuleType.NotContains;

                case ExcelConditionalFormattingConstants.Operators.NotEqual:
                    return eExcelConditionalFormattingRuleType.NotEqual;
                default:
                    throw new Exception(
                      ExcelConditionalFormattingConstants.Errors.UnexistentOperatorTypeAttribute);
            }
        }
        private static eExcelConditionalFormattingRuleType GetIconSetType(XmlNode topNode, XmlNamespaceManager nameSpaceManager)
        {
            var node = topNode.SelectSingleNode("d:iconSet/@iconSet", nameSpaceManager);
            if (node == null)
            {
                return eExcelConditionalFormattingRuleType.ThreeIconSet;
            }

            var v = node.Value;

            return v[0] switch
            {
                '3' => eExcelConditionalFormattingRuleType.ThreeIconSet,
                '4' => eExcelConditionalFormattingRuleType.FourIconSet,
                _ => eExcelConditionalFormattingRuleType.FiveIconSet
            };
        }

        /// <summary>
        /// Get the "colorScale" rule type according to the number of "cfvo" and "color" nodes.
        /// If we have excatly 2 "cfvo" and "color" childs, then we return "twoColorScale"
        /// </summary>
        /// <returns>TwoColorScale or ThreeColorScale</returns>
        internal static eExcelConditionalFormattingRuleType GetColorScaleType(
          XmlNode topNode,
          XmlNamespaceManager nameSpaceManager)
        {
            // Get the <cfvo> nodes
            var cfvoNodes = topNode.SelectNodes(
              string.Format(
                "{0}/{1}",
                ExcelConditionalFormattingConstants.Paths.ColorScale,
                ExcelConditionalFormattingConstants.Paths.Cfvo),
              nameSpaceManager);

            // Get the <color> nodes
            var colorNodes = topNode.SelectNodes(
              string.Format(
                "{0}/{1}",
                ExcelConditionalFormattingConstants.Paths.ColorScale,
                ExcelConditionalFormattingConstants.Paths.Color),
              nameSpaceManager);

            // We determine if it is "TwoColorScale" or "ThreeColorScale" by the
            // number of <cfvo> and <color> inside the <colorScale> node
            if (cfvoNodes == null || cfvoNodes.Count < 2 || cfvoNodes.Count > 3
              || colorNodes == null || colorNodes.Count < 2 || colorNodes.Count > 3
              || cfvoNodes.Count != colorNodes.Count)
            {
                throw new Exception(
                  ExcelConditionalFormattingConstants.Errors.WrongNumberCfvoColorNodes);
            }

            // Return the corresponding rule type (TwoColorScale or ThreeColorScale)
            return cfvoNodes.Count == 2
              ? eExcelConditionalFormattingRuleType.TwoColorScale
              : eExcelConditionalFormattingRuleType.ThreeColorScale;
        }

        /// <summary>
        /// Get the "aboveAverage" rule type according to the follwoing attributes:
        /// "AboveAverage", "EqualAverage" and "StdDev".
        /// 
        /// @StdDev greater than "0"                              == AboveStdDev
        /// @StdDev less than "0"                                 == BelowStdDev
        /// @AboveAverage = "1"/null and @EqualAverage = "0"/null == AboveAverage
        /// @AboveAverage = "1"/null and @EqualAverage = "1"      == AboveOrEqualAverage
        /// @AboveAverage = "0" and @EqualAverage = "0"/null      == BelowAverage
        /// @AboveAverage = "0" and @EqualAverage = "1"           == BelowOrEqualAverage
        /// /// </summary>
        /// <returns>AboveAverage, AboveOrEqualAverage, BelowAverage or BelowOrEqualAverage</returns>
        internal static eExcelConditionalFormattingRuleType GetAboveAverageType(
          XmlNode topNode,
          XmlNamespaceManager nameSpaceManager)
        {
            // Get @StdDev attribute
            var stdDev = ExcelConditionalFormattingHelper.GetAttributeIntNullable(
              topNode,
              ExcelConditionalFormattingConstants.Attributes.StdDev);

            switch (stdDev)
            {
                case > 0:
                    // @StdDev > "0" --> AboveStdDev
                    return eExcelConditionalFormattingRuleType.AboveStdDev;
                case < 0:
                    // @StdDev < "0" --> BelowStdDev
                    return eExcelConditionalFormattingRuleType.BelowStdDev;
            }

            // Get @AboveAverage attribute
            var isAboveAverage = ExcelConditionalFormattingHelper.GetAttributeBoolNullable(
              topNode,
              ExcelConditionalFormattingConstants.Attributes.AboveAverage);

            // Get @EqualAverage attribute
            var isEqualAverage = ExcelConditionalFormattingHelper.GetAttributeBoolNullable(
              topNode,
              ExcelConditionalFormattingConstants.Attributes.EqualAverage);

            if (isAboveAverage is null or true)
            {
                return isEqualAverage == true ?
                    // @AboveAverage = "1"/null and @EqualAverage = "1" == AboveOrEqualAverage
                    eExcelConditionalFormattingRuleType.AboveOrEqualAverage :
                    // @AboveAverage = "1"/null and @EqualAverage = "0"/null == AboveAverage
                    eExcelConditionalFormattingRuleType.AboveAverage;
            }

            return isEqualAverage == true ?
                // @AboveAverage = "0" and @EqualAverage = "1" == BelowOrEqualAverage
                eExcelConditionalFormattingRuleType.BelowOrEqualAverage :
                // @AboveAverage = "0" and @EqualAverage = "0"/null == BelowAverage
                eExcelConditionalFormattingRuleType.BelowAverage;
        }

        /// <summary>
        /// Get the "top10" rule type according to the follwoing attributes:
        /// "Bottom" and "Percent"
        /// 
        /// @Bottom = "1" and @Percent = "0"/null       == Bottom
        /// @Bottom = "1" and @Percent = "1"            == BottomPercent
        /// @Bottom = "0"/null and @Percent = "0"/null  == Top
        /// @Bottom = "0"/null and @Percent = "1"       == TopPercent
        /// /// </summary>
        /// <returns>Top, TopPercent, Bottom or BottomPercent</returns>
        public static eExcelConditionalFormattingRuleType GetTop10Type(
          XmlNode topNode,
          XmlNamespaceManager nameSpaceManager)
        {
            // Get @Bottom attribute
            var isBottom = ExcelConditionalFormattingHelper.GetAttributeBoolNullable(
              topNode,
              ExcelConditionalFormattingConstants.Attributes.Bottom);

            // Get @Percent attribute
            var isPercent = ExcelConditionalFormattingHelper.GetAttributeBoolNullable(
              topNode,
              ExcelConditionalFormattingConstants.Attributes.Percent);

            if (isBottom == true)
            {
                return isPercent == true ?
                    // @Bottom = "1" and @Percent = "1" == BottomPercent
                    eExcelConditionalFormattingRuleType.BottomPercent :
                    // @Bottom = "1" and @Percent = "0"/null == Bottom
                    eExcelConditionalFormattingRuleType.Bottom;
            }

            return isPercent == true ?
                // @Bottom = "0"/null and @Percent = "1" == TopPercent
                eExcelConditionalFormattingRuleType.TopPercent :
                // @Bottom = "0"/null and @Percent = "0"/null == Top
                eExcelConditionalFormattingRuleType.Top;
        }

        /// <summary>
        /// Get the "timePeriod" rule type according to "TimePeriod" attribute.
        /// /// </summary>
        /// <returns>Last7Days, LastMonth etc.</returns>
        public static eExcelConditionalFormattingRuleType GetTimePeriodType(
          XmlNode topNode,
          XmlNamespaceManager nameSpaceManager)
        {
            var timePeriod = ExcelConditionalFormattingTimePeriodType.GetTypeByAttribute(
              ExcelConditionalFormattingHelper.GetAttributeString(
                topNode,
                ExcelConditionalFormattingConstants.Attributes.TimePeriod));

            return timePeriod switch
            {
                eExcelConditionalFormattingTimePeriodType.Last7Days => eExcelConditionalFormattingRuleType.Last7Days,
                eExcelConditionalFormattingTimePeriodType.LastMonth => eExcelConditionalFormattingRuleType.LastMonth,
                eExcelConditionalFormattingTimePeriodType.LastWeek => eExcelConditionalFormattingRuleType.LastWeek,
                eExcelConditionalFormattingTimePeriodType.NextMonth => eExcelConditionalFormattingRuleType.NextMonth,
                eExcelConditionalFormattingTimePeriodType.NextWeek => eExcelConditionalFormattingRuleType.NextWeek,
                eExcelConditionalFormattingTimePeriodType.ThisMonth => eExcelConditionalFormattingRuleType.ThisMonth,
                eExcelConditionalFormattingTimePeriodType.ThisWeek => eExcelConditionalFormattingRuleType.ThisWeek,
                eExcelConditionalFormattingTimePeriodType.Today => eExcelConditionalFormattingRuleType.Today,
                eExcelConditionalFormattingTimePeriodType.Tomorrow => eExcelConditionalFormattingRuleType.Tomorrow,
                eExcelConditionalFormattingTimePeriodType.Yesterday => eExcelConditionalFormattingRuleType.Yesterday,
                _ => throw new Exception(ExcelConditionalFormattingConstants.Errors.UnexistentTimePeriodTypeAttribute)
            };
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static string GetAttributeByType(
          eExcelConditionalFormattingRuleType type)
        {
            switch (type)
            {
                case eExcelConditionalFormattingRuleType.AboveAverage:
                case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
                case eExcelConditionalFormattingRuleType.BelowAverage:
                case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
                case eExcelConditionalFormattingRuleType.AboveStdDev:
                case eExcelConditionalFormattingRuleType.BelowStdDev:
                    return ExcelConditionalFormattingConstants.RuleType.AboveAverage;

                case eExcelConditionalFormattingRuleType.Bottom:
                case eExcelConditionalFormattingRuleType.BottomPercent:
                case eExcelConditionalFormattingRuleType.Top:
                case eExcelConditionalFormattingRuleType.TopPercent:
                    return ExcelConditionalFormattingConstants.RuleType.Top10;

                case eExcelConditionalFormattingRuleType.Last7Days:
                case eExcelConditionalFormattingRuleType.LastMonth:
                case eExcelConditionalFormattingRuleType.LastWeek:
                case eExcelConditionalFormattingRuleType.NextMonth:
                case eExcelConditionalFormattingRuleType.NextWeek:
                case eExcelConditionalFormattingRuleType.ThisMonth:
                case eExcelConditionalFormattingRuleType.ThisWeek:
                case eExcelConditionalFormattingRuleType.Today:
                case eExcelConditionalFormattingRuleType.Tomorrow:
                case eExcelConditionalFormattingRuleType.Yesterday:
                    return ExcelConditionalFormattingConstants.RuleType.TimePeriod;

                case eExcelConditionalFormattingRuleType.Between:
                case eExcelConditionalFormattingRuleType.Equal:
                case eExcelConditionalFormattingRuleType.GreaterThan:
                case eExcelConditionalFormattingRuleType.GreaterThanOrEqual:
                case eExcelConditionalFormattingRuleType.LessThan:
                case eExcelConditionalFormattingRuleType.LessThanOrEqual:
                case eExcelConditionalFormattingRuleType.NotBetween:
                case eExcelConditionalFormattingRuleType.NotEqual:
                    return ExcelConditionalFormattingConstants.RuleType.CellIs;

                case eExcelConditionalFormattingRuleType.ThreeIconSet:
                case eExcelConditionalFormattingRuleType.FourIconSet:
                case eExcelConditionalFormattingRuleType.FiveIconSet:
                    return ExcelConditionalFormattingConstants.RuleType.IconSet;

                case eExcelConditionalFormattingRuleType.ThreeColorScale:
                case eExcelConditionalFormattingRuleType.TwoColorScale:
                    return ExcelConditionalFormattingConstants.RuleType.ColorScale;

                case eExcelConditionalFormattingRuleType.BeginsWith:
                    return ExcelConditionalFormattingConstants.RuleType.BeginsWith;

                case eExcelConditionalFormattingRuleType.ContainsBlanks:
                    return ExcelConditionalFormattingConstants.RuleType.ContainsBlanks;

                case eExcelConditionalFormattingRuleType.ContainsErrors:
                    return ExcelConditionalFormattingConstants.RuleType.ContainsErrors;

                case eExcelConditionalFormattingRuleType.ContainsText:
                    return ExcelConditionalFormattingConstants.RuleType.ContainsText;

                case eExcelConditionalFormattingRuleType.DuplicateValues:
                    return ExcelConditionalFormattingConstants.RuleType.DuplicateValues;

                case eExcelConditionalFormattingRuleType.EndsWith:
                    return ExcelConditionalFormattingConstants.RuleType.EndsWith;

                case eExcelConditionalFormattingRuleType.Expression:
                    return ExcelConditionalFormattingConstants.RuleType.Expression;

                case eExcelConditionalFormattingRuleType.NotContainsBlanks:
                    return ExcelConditionalFormattingConstants.RuleType.NotContainsBlanks;

                case eExcelConditionalFormattingRuleType.NotContainsErrors:
                    return ExcelConditionalFormattingConstants.RuleType.NotContainsErrors;

                case eExcelConditionalFormattingRuleType.NotContainsText:
                    return ExcelConditionalFormattingConstants.RuleType.NotContainsText;

                case eExcelConditionalFormattingRuleType.UniqueValues:
                    return ExcelConditionalFormattingConstants.RuleType.UniqueValues;

                case eExcelConditionalFormattingRuleType.DataBar:
                    return ExcelConditionalFormattingConstants.RuleType.DataBar;
            }

            throw new Exception(
              ExcelConditionalFormattingConstants.Errors.MissingRuleType);
        }

        /// <summary>
        /// Return cfvo §18.3.1.11 parent according to the rule type
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static string GetCfvoParentPathByType(
          eExcelConditionalFormattingRuleType type)
        {
            switch (type)
            {
                case eExcelConditionalFormattingRuleType.TwoColorScale:
                case eExcelConditionalFormattingRuleType.ThreeColorScale:
                    return ExcelConditionalFormattingConstants.Paths.ColorScale;

                case eExcelConditionalFormattingRuleType.ThreeIconSet:
                case eExcelConditionalFormattingRuleType.FourIconSet:
                case eExcelConditionalFormattingRuleType.FiveIconSet:
                    return ExcelConditionalFormattingConstants.RuleType.IconSet;

                case eExcelConditionalFormattingRuleType.DataBar:
                    return ExcelConditionalFormattingConstants.RuleType.DataBar;
            }

            throw new Exception(
              ExcelConditionalFormattingConstants.Errors.MissingRuleType);
        }
    }
}