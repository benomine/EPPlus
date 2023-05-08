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
 * Jan Källman		    Initial Release		       2009-10-01
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;

namespace OfficeOpenXml
{
    /// <summary>
    /// The collection of worksheets for the workbook
    /// </summary>
    public class ExcelWorksheets : XmlHelper, IEnumerable<ExcelWorksheet>, IDisposable
    {
        #region Private Properties
        private ExcelPackage _pck;
        private Dictionary<int, ExcelWorksheet> _worksheets;
        private XmlNamespaceManager _namespaceManager;
        #endregion
        #region ExcelWorksheets Constructor
        internal ExcelWorksheets(ExcelPackage pck, XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            _pck = pck;
            _namespaceManager = nsm;
            _worksheets = new Dictionary<int, ExcelWorksheet>();
            var positionID = _pck._worksheetAdd;

            foreach (XmlNode sheetNode in topNode.ChildNodes)
            {
                if (sheetNode.NodeType == XmlNodeType.Element)
                {
                    var name = sheetNode.Attributes["name"].Value;
                    //Get the relationship id
                    var relId = sheetNode.Attributes.GetNamedItem("id", ExcelPackage.schemaRelationships).Value;
                    var sheetID = Convert.ToInt32(sheetNode.Attributes["sheetId"].Value);

                    //Hidden property
                    var hidden = eWorkSheetHidden.Visible;
                    XmlNode attr = sheetNode.Attributes["state"];
                    if (attr != null)
                        hidden = TranslateHidden(attr.Value);

                    var sheetRelation = pck.Workbook.Part.GetRelationship(relId);
                    var uriWorksheet = UriHelper.ResolvePartUri(pck.Workbook.WorkbookUri, sheetRelation.TargetUri);

                    //add the worksheet
                    if (sheetRelation.RelationshipType.EndsWith("chartsheet"))
                    {
                        _worksheets.Add(positionID, new ExcelChartsheet(_namespaceManager, _pck, relId, uriWorksheet, name, sheetID, positionID, hidden));
                    }
                    else
                    {
                        _worksheets.Add(positionID, new ExcelWorksheet(_namespaceManager, _pck, relId, uriWorksheet, name, sheetID, positionID, hidden));
                    }
                    positionID++;
                }
            }
        }

        private eWorkSheetHidden TranslateHidden(string value)
        {
            return value switch
            {
                "hidden" => eWorkSheetHidden.Hidden,
                "veryHidden" => eWorkSheetHidden.VeryHidden,
                _ => eWorkSheetHidden.Visible
            };
        }
        #endregion
        #region ExcelWorksheets Public Properties
        /// <summary>
        /// Returns the number of worksheets in the workbook
        /// </summary>
        public int Count => _worksheets.Count;

        #endregion
        private const string ERR_DUP_WORKSHEET = "A worksheet with this name already exists in the workbook";
        internal const string WORKSHEET_CONTENTTYPE = @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        internal const string CHARTSHEET_CONTENTTYPE = @"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
        #region ExcelWorksheets Public Methods
        /// <summary>
        /// Foreach support
        /// </summary>
        /// <returns>An enumerator</returns>
        public IEnumerator<ExcelWorksheet> GetEnumerator()
        {
            return _worksheets.Values.GetEnumerator();
        }
        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _worksheets.Values.GetEnumerator();
        }

        #endregion
        #region Add Worksheet
        /// <summary>
        /// Adds a new blank worksheet.
        /// </summary>
        /// <param name="name">The name of the workbook</param>
        public ExcelWorksheet Add(string name)
        {
            var worksheet = AddSheet(name, false, null);
            return worksheet;
        }
        private ExcelWorksheet AddSheet(string Name, bool isChart, eChartType? chartType, ExcelPivotTable pivotTableSource = null)
        {
            lock (_worksheets)
            {
                Name = ValidateFixSheetName(Name);
                if (GetByName(Name) != null)
                {
                    throw new InvalidOperationException(ERR_DUP_WORKSHEET + " : " + Name);
                }

                GetSheetURI(ref Name, out var sheetID, out var uriWorksheet, isChart);
                var worksheetPart = _pck.Package.CreatePart(uriWorksheet, isChart ? CHARTSHEET_CONTENTTYPE : WORKSHEET_CONTENTTYPE, _pck.Compression);

                //Create the new, empty worksheet and save it to the package
                var streamWorksheet = new StreamWriter(worksheetPart.GetStream(FileMode.Create, FileAccess.Write));
                var worksheetXml = CreateNewWorksheet(isChart);
                worksheetXml.Save(streamWorksheet);
                _pck.Package.Flush();

                var rel = CreateWorkbookRel(Name, sheetID, uriWorksheet, isChart);

                var positionID = _worksheets.Count + _pck._worksheetAdd;
                ExcelWorksheet worksheet;
                if (isChart)
                {
                    worksheet = new ExcelChartsheet(_namespaceManager, _pck, rel, uriWorksheet, Name, sheetID, positionID, eWorkSheetHidden.Visible, (eChartType)chartType, pivotTableSource);
                }
                else
                {
                    worksheet = new ExcelWorksheet(_namespaceManager, _pck, rel, uriWorksheet, Name, sheetID, positionID, eWorkSheetHidden.Visible);
                }

                _worksheets.Add(positionID, worksheet);
                if (_pck.Workbook.VbaProject != null)
                {
                    var name = _pck.Workbook.VbaProject.GetModuleNameFromWorksheet(worksheet);
                    _pck.Workbook.VbaProject.Modules.Add(new ExcelVBAModule(worksheet.CodeNameChange) { Name = name, Code = "", Attributes = _pck.Workbook.VbaProject.GetDocumentAttributes(Name, "0{00020820-0000-0000-C000-000000000046}"), Type = eModuleType.Document, HelpContext = 0 });
                    worksheet.CodeModuleName = name;

                }
                return worksheet;
            }
        }
        /// <summary>
        /// Adds a copy of a worksheet
        /// </summary>
        /// <param name="Name">The name of the workbook</param>
        /// <param name="copy">The worksheet to be copied</param>
        public ExcelWorksheet Add(string Name, ExcelWorksheet copy)
        {
            lock (_worksheets)
            {
                int sheetID;
                Uri uriWorksheet;
                if (copy is ExcelChartsheet)
                {
                    throw new ArgumentException("Can not copy a chartsheet");
                }
                if (GetByName(Name) != null)
                {
                    throw new InvalidOperationException(ERR_DUP_WORKSHEET);
                }

                GetSheetURI(ref Name, out sheetID, out uriWorksheet, false);

                //Create a copy of the worksheet XML
                var worksheetPart = _pck.Package.CreatePart(uriWorksheet, WORKSHEET_CONTENTTYPE, _pck.Compression);
                var streamWorksheet = new StreamWriter(worksheetPart.GetStream(FileMode.Create, FileAccess.Write));
                var worksheetXml = new XmlDocument();
                worksheetXml.LoadXml(copy.WorksheetXml.OuterXml);
                worksheetXml.Save(streamWorksheet);
                _pck.Package.Flush();


                //Create a relation to the workbook
                var relID = CreateWorkbookRel(Name, sheetID, uriWorksheet, false);
                var added = new ExcelWorksheet(_namespaceManager, _pck, relID, uriWorksheet, Name, sheetID, _worksheets.Count + _pck._worksheetAdd, eWorkSheetHidden.Visible);

                //Copy comments
                if (copy.Comments.Count > 0)
                {
                    CopyComment(copy, added);
                }
                else if (copy.VmlDrawingsComments.Count > 0)    //Vml drawings are copied as part of the comments. 
                {
                    CopyVmlDrawing(copy, added);
                }

                //Copy HeaderFooter
                CopyHeaderFooterPictures(copy, added);

                //Copy all relationships 
                //CopyRelationShips(Copy, added);
                if (copy.Drawings.Count > 0)
                {
                    CopyDrawing(copy, added);
                }
                if (copy.Tables.Count > 0)
                {
                    CopyTable(copy, added);
                }
                if (copy.PivotTables.Count > 0)
                {
                    CopyPivotTable(copy, added);
                }
                if (copy.Names.Count > 0)
                {
                    CopySheetNames(copy, added);
                }

                //Copy all cells
                CloneCells(copy, added);

                //Copy the VBA code
                if (_pck.Workbook.VbaProject != null)
                {
                    var name = _pck.Workbook.VbaProject.GetModuleNameFromWorksheet(added);
                    _pck.Workbook.VbaProject.Modules.Add(new ExcelVBAModule(added.CodeNameChange) { Name = name, Code = copy.CodeModule.Code, Attributes = _pck.Workbook.VbaProject.GetDocumentAttributes(Name, "0{00020820-0000-0000-C000-000000000046}"), Type = eModuleType.Document, HelpContext = 0 });
                    copy.CodeModuleName = name;
                }

                _worksheets.Add(_worksheets.Count + _pck._worksheetAdd, added);

                var pageSetup = added.WorksheetXml.SelectSingleNode("//d:pageSetup", _namespaceManager);
                var attr = (XmlAttribute)pageSetup?.Attributes.GetNamedItem("id", ExcelPackage.schemaRelationships);
                if (attr == null) return added;
                relID = attr.Value;
                pageSetup.Attributes.Remove(attr);
                return added;
            }
        }
        /// <summary>
        /// Adds a chartsheet to the workbook.
        /// </summary>
        /// <param name="name">The name of the worksheet</param>
        /// <param name="chartType">The type of chart</param>
        /// <returns></returns>
        public ExcelChartsheet AddChart(string name, eChartType chartType)
        {
            return (ExcelChartsheet)AddSheet(name, true, chartType, null);
        }
        /// <summary>
        /// Adds a chartsheet to the workbook.
        /// </summary>
        /// <param name="name">The name of the worksheet</param>
        /// <param name="chartType">The type of chart</param>
        /// <param name="pivotTableSource">The pivottable source</param>
        /// <returns></returns>
        public ExcelChartsheet AddChart(string name, eChartType chartType, ExcelPivotTable pivotTableSource)
        {
            return (ExcelChartsheet)AddSheet(name, true, chartType, pivotTableSource);
        }
        private void CopySheetNames(ExcelWorksheet copy, ExcelWorksheet added)
        {
            foreach (var name in copy.Names)
            {
                ExcelNamedRange newName;
                if (!name.IsName)
                {
                    newName = added.Names.Add(name.Name, name.WorkSheet == copy.Name ? added.Cells[name.FirstAddress] : added.Workbook.Worksheets[name.WorkSheet].Cells[name.FirstAddress]);
                }
                else if (!string.IsNullOrEmpty(name.NameFormula))
                {
                    newName=added.Names.AddFormula(name.Name, name.Formula);
                }
                else
                {
                    newName=added.Names.AddValue(name.Name, name.Value);
                }
                newName.NameComment = name.NameComment;
            }
        }

        private void CopyTable(ExcelWorksheet copy, ExcelWorksheet added)
        {
            var prevName = "";
            foreach (var tbl in copy.Tables)
            {
                var xml = tbl.TableXml.OuterXml;
                string name;
                if (prevName == "")
                {
                    name = copy.Tables.GetNewTableName();
                }
                else
                {
                    var ix = int.Parse(prevName.Substring(5)) + 1;
                    name = string.Format("Table{0}", ix);
                    while (_pck.Workbook.ExistsPivotTableName(name))
                    {
                        name = string.Format("Table{0}", ++ix);
                    }
                }
                _pck.Workbook.ReadAllTables();

                var Id = _pck.Workbook._nextTableID++;
                prevName = name;
                var xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xml);
                xmlDoc.SelectSingleNode("//d:table/@id", tbl.NameSpaceManager).Value = Id.ToString();
                xmlDoc.SelectSingleNode("//d:table/@name", tbl.NameSpaceManager).Value = name;
                xmlDoc.SelectSingleNode("//d:table/@displayName", tbl.NameSpaceManager).Value = name;
                xml = xmlDoc.OuterXml;

                //var uriTbl = new Uri(string.Format("/xl/tables/table{0}.xml", Id), UriKind.Relative);
                var uriTbl = GetNewUri(_pck.Package, "/xl/tables/table{0}.xml", ref Id);
                if (_pck.Workbook._nextTableID < Id) _pck.Workbook._nextTableID = Id;

                var part = _pck.Package.CreatePart(uriTbl, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", _pck.Compression);
                var streamTbl = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                streamTbl.Write(xml);
                //streamTbl.Close();
                streamTbl.Flush();

                //create the relationship and add the ID to the worksheet xml.
                var rel = added.Part.CreateRelationship(UriHelper.GetRelativeUri(added.WorksheetUri, uriTbl), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/table");

                if (tbl.RelationshipID == null)
                {
                    var topNode = added.WorksheetXml.SelectSingleNode("//d:tableParts", tbl.NameSpaceManager);
                    if (topNode == null)
                    {
                        added.CreateNode("d:tableParts");
                        topNode = added.WorksheetXml.SelectSingleNode("//d:tableParts", tbl.NameSpaceManager);
                    }
                    var elem = added.WorksheetXml.CreateElement("tablePart", ExcelPackage.schemaMain);
                    topNode.AppendChild(elem);
                    elem.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);
                }
                else
                {
                    var relAtt = added.WorksheetXml.SelectSingleNode(string.Format("//d:tableParts/d:tablePart/@r:id[.='{0}']", tbl.RelationshipID), tbl.NameSpaceManager) as XmlAttribute;
                    relAtt.Value = rel.Id;
                }
            }
        }
        private void CopyPivotTable(ExcelWorksheet copy, ExcelWorksheet added)
        {
            var prevName = "";
            foreach (var tbl in copy.PivotTables)
            {
                var xml = tbl.PivotTableXml.OuterXml;

                string name;
                if (copy.Workbook==added.Workbook || added.PivotTables._pivotTableNames.ContainsKey(tbl.Name))
                {
                    if (prevName == "")
                    {
                        name = added.PivotTables.GetNewTableName();
                    }
                    else
                    {
                        var ix = int.Parse(prevName.Substring(10)) + 1;
                        name = string.Format("PivotTable{0}", ix);
                        while (_pck.Workbook.ExistsPivotTableName(name))
                        {
                            name = string.Format("PivotTable{0}", ++ix);
                        }
                    }
                }
                else
                {
                    name = tbl.Name;
                }
                prevName=name;
                var xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xml);
                xmlDoc.SelectSingleNode("//d:pivotTableDefinition/@name", tbl.NameSpaceManager).Value = name;
                var cacheId = tbl.CacheID;
                if (!added.Workbook.ExistsPivotCache(tbl.CacheID, ref cacheId))
                {
                    xmlDoc.SelectSingleNode("//d:pivotTableDefinition/@cacheId", tbl.NameSpaceManager).Value = cacheId.ToString();

                }
                xml = xmlDoc.OuterXml;

                var Id = _pck.Workbook._nextPivotTableID++;
                var uriTbl = GetNewUri(_pck.Package, "/xl/pivotTables/pivotTable{0}.xml", ref Id);
                if (_pck.Workbook._nextPivotTableID < Id) _pck.Workbook._nextPivotTableID = Id;
                var partTbl = _pck.Package.CreatePart(uriTbl, ExcelPackage.schemaPivotTable, _pck.Compression);
                var streamTbl = new StreamWriter(partTbl.GetStream(FileMode.Create, FileAccess.Write));
                streamTbl.Write(xml);
                streamTbl.Flush();

                xml = tbl.CacheDefinition.CacheDefinitionXml.OuterXml;
                var uriCd = GetNewUri(_pck.Package, "/xl/pivotCache/pivotcachedefinition{0}.xml", ref Id);
                var partCd = _pck.Package.CreatePart(uriCd, ExcelPackage.schemaPivotCacheDefinition, _pck.Compression);
                var streamCd = new StreamWriter(partCd.GetStream(FileMode.Create, FileAccess.Write));
                streamCd.Write(xml);
                streamCd.Flush();

                added.Workbook.AddPivotTable(cacheId.ToString(), uriCd);

                xml = "<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" count=\"0\" />";
                var uriRec = new Uri(string.Format("/xl/pivotCache/pivotCacheRecords{0}.xml", Id), UriKind.Relative);
                while (_pck.Package.PartExists(uriRec))
                {
                    uriRec = new Uri(string.Format("/xl/pivotCache/pivotCacheRecords{0}.xml", ++Id), UriKind.Relative);
                }
                var partRec = _pck.Package.CreatePart(uriRec, ExcelPackage.schemaPivotCacheRecords, _pck.Compression);
                var streamRec = new StreamWriter(partRec.GetStream(FileMode.Create, FileAccess.Write));
                streamRec.Write(xml);
                streamRec.Flush();

                added.Part.CreateRelationship(UriHelper.ResolvePartUri(added.WorksheetUri, uriTbl), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotTable");
                partTbl.CreateRelationship(UriHelper.ResolvePartUri(tbl.Relationship.SourceUri, uriCd), tbl.CacheDefinition.Relationship.TargetMode, tbl.CacheDefinition.Relationship.RelationshipType);
                partCd.CreateRelationship(UriHelper.ResolvePartUri(uriCd, uriRec), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheRecords");

            }
            added._pivotTables = null;
        }
        private void CopyHeaderFooterPictures(ExcelWorksheet copy, ExcelWorksheet added)
        {
            if (copy.TopNode != null && copy.TopNode.SelectSingleNode("d:headerFooter", NameSpaceManager)==null) return;
            //Copy the texts
            if (copy.HeaderFooter._oddHeader!=null) CopyText(copy.HeaderFooter._oddHeader, added.HeaderFooter.OddHeader);
            if (copy.HeaderFooter._oddFooter != null) CopyText(copy.HeaderFooter._oddFooter, added.HeaderFooter.OddFooter);
            if (copy.HeaderFooter._evenHeader != null) CopyText(copy.HeaderFooter._evenHeader, added.HeaderFooter.EvenHeader);
            if (copy.HeaderFooter._evenFooter != null) CopyText(copy.HeaderFooter._evenFooter, added.HeaderFooter.EvenFooter);
            if (copy.HeaderFooter._firstHeader != null) CopyText(copy.HeaderFooter._firstHeader, added.HeaderFooter.FirstHeader);
            if (copy.HeaderFooter._firstFooter != null) CopyText(copy.HeaderFooter._firstFooter, added.HeaderFooter.FirstFooter);

            //Copy any images;
            if (copy.HeaderFooter.Pictures.Count > 0)
            {
                var source = copy.HeaderFooter.Pictures.Uri;
                var dest = GetNewUri(_pck.Package, @"/xl/drawings/vmlDrawing{0}.vml");
                added.DeleteNode("d:legacyDrawingHF");

                foreach (ExcelVmlDrawingPicture pic in copy.HeaderFooter.Pictures)
                {
                    var item = added.HeaderFooter.Pictures.Add(pic.Id, pic.ImageUri, pic.Title, pic.Width, pic.Height);
                    foreach (XmlAttribute att in pic.TopNode.Attributes)
                    {
                        (item.TopNode as XmlElement).SetAttribute(att.Name, att.Value);
                    }
                    item.TopNode.InnerXml = pic.TopNode.InnerXml;
                }
            }
        }

        private void CopyText(ExcelHeaderFooterText from, ExcelHeaderFooterText to)
        {
            to.LeftAlignedText=from.LeftAlignedText;
            to.CenteredText = from.CenteredText;
            to.RightAlignedText = from.RightAlignedText;
        }
        private void CloneCells(ExcelWorksheet copy, ExcelWorksheet added)
        {
            var sameWorkbook = copy.Workbook == _pck.Workbook;

            var doAdjust = _pck.DoAdjustDrawings;
            _pck.DoAdjustDrawings = false;
            //Merged cells
            foreach (var r in copy.MergedCells)     //Issue #94
            {
                added.MergedCells.Add(new ExcelAddress(r), false);
            }

            //Shared Formulas   
            foreach (var key in copy._sharedFormulas.Keys)
            {
                added._sharedFormulas.Add(key, copy._sharedFormulas[key].Clone());
            }

            var styleCashe = new Dictionary<int, int>();
            //Cells
            var val = new CellsStoreEnumerator<ExcelCoreValue>(copy._values);
            while (val.Next())
            {
                var row = val.Row;
                var col = val.Column;
                var styleID = 0;
                if (row == 0)
                {
                    if (copy.GetValueInner(row, col) is ExcelColumn c)
                    {
                        var clone = c.Clone(added, c.ColumnMin);
                        clone.StyleID = c.StyleID;
                        added.SetValueInner(row, col, clone);
                        styleID = c.StyleID;
                    }
                }
                else if (col == 0)
                {
                    var r = copy.Row(row);
                    if (r != null)
                    {
                        r.Clone(added);
                        styleID = r.StyleID;
                    }

                }
                else
                {
                    styleID = CopyValues(copy, added, row, col);
                }

                if (sameWorkbook) continue;
                if (styleCashe.ContainsKey(styleID))
                {
                    added.SetStyleInner(row, col, styleCashe[styleID]);
                }
                else
                {
                    var s = added.Workbook.Styles.CloneStyle(copy.Workbook.Styles, styleID);
                    styleCashe.Add(styleID, s);
                    added.SetStyleInner(row, col, s);
                }
            }
            added._package.DoAdjustDrawings = doAdjust;
        }

        private int CopyValues(ExcelWorksheet copy, ExcelWorksheet added, int row, int col)
        {
            added.SetValueInner(row, col, copy.GetValueInner(row, col));
            byte fl = 0;
            if (copy._flags.Exists(row, col, ref fl))
            {
                added._flags.SetValue(row, col, fl);
            }

            var v = copy._formulas.GetValue(row, col);
            if (v != null)
            {
                added.SetFormula(row, col, v);
            }
            var s = copy.GetStyleInner(row, col);
            if (s != 0)
            {
                added.SetStyleInner(row, col, s);
            }
            var f = copy._formulas.GetValue(row, col);
            if (f != null)
            {
                added._formulas.SetValue(row, col, f);
            }
            return s;
        }

        private void CopyComment(ExcelWorksheet copy, ExcelWorksheet workSheet)
        {
            var xml = copy.Comments.CommentXml.InnerXml;
            var uriComment = new Uri(string.Format("/xl/comments{0}.xml", workSheet.SheetID), UriKind.Relative);
            if (_pck.Package.PartExists(uriComment))
            {
                uriComment = GetNewUri(_pck.Package, "/xl/drawings/vmldrawing{0}.vml");
            }

            var part = _pck.Package.CreatePart(uriComment, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", _pck.Compression);

            var streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
            streamDrawing.Write(xml);
            streamDrawing.Flush();

            var commentRelation = workSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uriComment), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/comments");

            xml = copy.VmlDrawingsComments.VmlDrawingXml.InnerXml;

            var uriVml = new Uri(string.Format("/xl/drawings/vmldrawing{0}.vml", workSheet.SheetID), UriKind.Relative);
            if (_pck.Package.PartExists(uriVml))
            {
                uriVml = GetNewUri(_pck.Package, "/xl/drawings/vmldrawing{0}.vml");
            }

            var vmlPart = _pck.Package.CreatePart(uriVml, "application/vnd.openxmlformats-officedocument.vmlDrawing", _pck.Compression);
            var streamVml = new StreamWriter(vmlPart.GetStream(FileMode.Create, FileAccess.Write));
            streamVml.Write(xml);
            streamVml.Flush();

            var newVmlRel = workSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uriVml), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");

            var e = workSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _namespaceManager) as XmlElement;
            if (e == null)
            {
                workSheet.CreateNode("d:legacyDrawing");
                e = workSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _namespaceManager) as XmlElement;
            }

            e.SetAttribute("id", ExcelPackage.schemaRelationships, newVmlRel.Id);
        }
        private void CopyDrawing(ExcelWorksheet copy, ExcelWorksheet workSheet)
        {
            var xml = copy.Drawings.DrawingXml.OuterXml;
            var uriDraw = new Uri(string.Format("/xl/drawings/drawing{0}.xml", workSheet.SheetID), UriKind.Relative);
            var part = _pck.Package.CreatePart(uriDraw, "application/vnd.openxmlformats-officedocument.drawing+xml", _pck.Compression);
            var streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
            streamDrawing.Write(xml);
            streamDrawing.Flush();

            var drawXml = new XmlDocument();
            drawXml.LoadXml(xml);
            var drawRelation = workSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uriDraw), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/drawing");
            var e = workSheet.WorksheetXml.SelectSingleNode("//d:drawing", _namespaceManager) as XmlElement;
            e.SetAttribute("id", ExcelPackage.schemaRelationships, drawRelation.Id);

            for (var i = 0; i<copy.Drawings.Count; i++)
            {
                var draw = copy.Drawings[i];
                draw.AdjustPositionAndSize();
                if (draw is ExcelChart chart)
                {
                    xml = chart.ChartXml.InnerXml;

                    var uriChart = GetNewUri(_pck.Package, "/xl/charts/chart{0}.xml");
                    var chartPart = _pck.Package.CreatePart(uriChart, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", _pck.Compression);
                    var streamChart = new StreamWriter(chartPart.GetStream(FileMode.Create, FileAccess.Write));
                    streamChart.Write(xml);
                    streamChart.Flush();
                    var prevRelID = chart.TopNode.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart/@r:id", copy.Drawings.NameSpaceManager).Value;
                    var rel = part.CreateRelationship(UriHelper.GetRelativeUri(uriDraw, uriChart), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/chart");
                    var relAtt = drawXml.SelectSingleNode(string.Format("//c:chart/@r:id[.='{0}']", prevRelID), copy.Drawings.NameSpaceManager) as XmlAttribute;
                    relAtt.Value=rel.Id;
                }
                else if (draw is ExcelPicture picture)
                {
                    var uri = picture.UriPic;
                    if (!workSheet.Workbook._package.Package.PartExists(uri))
                    {
                        var picPart = workSheet.Workbook._package.Package.CreatePart(uri, picture.ContentType, CompressionLevel.None);
                        picture.Image.EncodedData.SaveTo(picPart.GetStream(FileMode.Create, FileAccess.Write));
                    }

                    var rel = part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");

                    var relAtt =
                        drawXml.SelectSingleNode(
                            string.Format(
                                "//xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name[.='{0}']/../../../xdr:blipFill/a:blip/@r:embed",
                                picture.Name), copy.Drawings.NameSpaceManager);
                    if (relAtt!=null)
                    {
                        relAtt.Value = rel.Id;
                    }
                    if (_pck._images.ContainsKey(picture.ImageHash))
                    {
                        _pck._images[picture.ImageHash].RefCount++;
                    }
                }
            }
            streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
            streamDrawing.Write(drawXml.OuterXml);
            streamDrawing.Flush();

            for (var i = 0; i<copy.Drawings.Count; i++)
            {
                var draw = copy.Drawings[i];
                var c = workSheet.Drawings[i];
                if (c == null) continue;
                c._left = draw._left;
                c._top = draw._top;
                c._height = draw._height;
                c._width = draw._width;
            }
        }

        private void CopyVmlDrawing(ExcelWorksheet origSheet, ExcelWorksheet newSheet)
        {
            var xml = origSheet.VmlDrawingsComments.VmlDrawingXml.OuterXml;
            var vmlUri = new Uri(string.Format("/xl/drawings/vmlDrawing{0}.vml", newSheet.SheetID), UriKind.Relative);
            var part = _pck.Package.CreatePart(vmlUri, "application/vnd.openxmlformats-officedocument.vmlDrawing", _pck.Compression);
            using (var streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write)))
            {
                streamDrawing.Write(xml);
                streamDrawing.Flush();
            }

            //Add the relationship ID to the worksheet xml.
            var vmlRelation = newSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(newSheet.WorksheetUri, vmlUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");
            var e = newSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _namespaceManager) as XmlElement ??
                    newSheet.WorksheetXml.CreateNode(XmlNodeType.Entity, "//d:legacyDrawing", _namespaceManager.LookupNamespace("d")) as XmlElement;

            e?.SetAttribute("id", ExcelPackage.schemaRelationships, vmlRelation.Id);
        }

        private string CreateWorkbookRel(string name, int sheetId, Uri uriWorksheet, bool isChart)
        {
            //Create the relationship between the workbook and the new worksheet
            var rel = _pck.Workbook.Part.CreateRelationship(UriHelper.GetRelativeUri(_pck.Workbook.WorkbookUri, uriWorksheet), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/" + (isChart ? "chartsheet" : "worksheet"));
            _pck.Package.Flush();

            //Create the new sheet node
            var worksheetNode = _pck.Workbook.WorkbookXml.CreateElement("sheet", ExcelPackage.schemaMain);
            worksheetNode.SetAttribute("name", name);
            worksheetNode.SetAttribute("sheetId", sheetId.ToString());
            worksheetNode.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);

            TopNode.AppendChild(worksheetNode);
            return rel.Id;
        }

        private void GetSheetURI(ref string name, out int sheetId, out Uri uriWorksheet, bool isChart)
        {
            name = ValidateFixSheetName(name);
            sheetId = this.Any() ? this.Max(ws => ws.SheetID) + 1 : 1;
            var uriId = sheetId;

            // get the next available worhsheet uri
            do
            {
                uriWorksheet = isChart ? new Uri("/xl/chartsheets/chartsheet" + uriId + ".xml", UriKind.Relative) : new Uri("/xl/worksheets/sheet" + uriId + ".xml", UriKind.Relative);

                uriId++;
            } while (_pck.Package.PartExists(uriWorksheet));
        }

        internal string ValidateFixSheetName(string name)
        {
            //remove invalid characters
            if (ValidateName(name))
            {
                if (name.IndexOf(':') > -1) name = name.Replace(":", " ");
                if (name.IndexOf('/') > -1) name = name.Replace("/", " ");
                if (name.IndexOf('\\') > -1) name = name.Replace("\\", " ");
                if (name.IndexOf('?') > -1) name = name.Replace("?", " ");
                if (name.IndexOf('[') > -1) name = name.Replace("[", " ");
                if (name.IndexOf(']') > -1) name = name.Replace("]", " ");
            }

            if (name.Trim() == "")
            {
                throw new ArgumentException("The worksheet can not have an empty name");
            }
            if (name.StartsWith("'") || name.EndsWith("'"))
            {
                throw new ArgumentException("The worksheet name can not start or end with an apostrophe.");
            }
            if (name.Length > 31) name = name.Substring(0, 31);   //A sheet can have max 31 char's            
            return name;
        }
        /// <summary>
        /// Validate the sheetname
        /// </summary>
        /// <param name="name">The Name</param>
        /// <returns>True if valid</returns>
        private bool ValidateName(string name)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(name, @":|\?|/|\\|\[|\]");
        }

        /// <summary>
        /// Creates the XML document representing a new empty worksheet
        /// </summary>
        /// <returns></returns>
        private XmlDocument CreateNewWorksheet(bool isChart)
        {
            var xmlDoc = new XmlDocument();
            var elemWs = xmlDoc.CreateElement(isChart ? "chartsheet" : "worksheet", ExcelPackage.schemaMain);
            elemWs.SetAttribute("xmlns:r", ExcelPackage.schemaRelationships);
            xmlDoc.AppendChild(elemWs);


            if (isChart)
            {
                var elemSheetPr = xmlDoc.CreateElement("sheetPr", ExcelPackage.schemaMain);
                elemWs.AppendChild(elemSheetPr);

                var elemSheetViews = xmlDoc.CreateElement("sheetViews", ExcelPackage.schemaMain);
                elemWs.AppendChild(elemSheetViews);

                var elemSheetView = xmlDoc.CreateElement("sheetView", ExcelPackage.schemaMain);
                elemSheetView.SetAttribute("workbookViewId", "0");
                elemSheetView.SetAttribute("zoomToFit", "1");

                elemSheetViews.AppendChild(elemSheetView);
            }
            else
            {
                var elemSheetViews = xmlDoc.CreateElement("sheetViews", ExcelPackage.schemaMain);
                elemWs.AppendChild(elemSheetViews);

                var elemSheetView = xmlDoc.CreateElement("sheetView", ExcelPackage.schemaMain);
                elemSheetView.SetAttribute("workbookViewId", "0");
                elemSheetViews.AppendChild(elemSheetView);

                var elemSheetFormatPr = xmlDoc.CreateElement("sheetFormatPr", ExcelPackage.schemaMain);
                elemSheetFormatPr.SetAttribute("defaultRowHeight", "15");
                elemWs.AppendChild(elemSheetFormatPr);

                var elemSheetData = xmlDoc.CreateElement("sheetData", ExcelPackage.schemaMain);
                elemWs.AppendChild(elemSheetData);
            }
            return xmlDoc;
        }
        #endregion
        #region Delete Worksheet
        /// <summary>
        /// Deletes a worksheet from the collection
        /// </summary>
        /// <param name="index">The position of the worksheet in the workbook</param>
        public void Delete(int index)
        {
            /*
            * Hack to prefetch all the drawings,
            * so that all the images are referenced, 
            * to prevent the deletion of the image file, 
            * when referenced more than once
            */
            foreach (var ws in _worksheets)
            {
                var drawings = ws.Value.Drawings;
            }

            var worksheet = _worksheets[index];
            if (worksheet.Drawings.Count > 0)
            {
                worksheet.Drawings.ClearDrawings();
            }

            //Remove all comments
            if (worksheet is not ExcelChartsheet && worksheet.Comments.Count > 0)
            {
                worksheet.Comments.Clear();
            }

            //Delete any parts still with relations to the Worksheet.
            DeleteRelationsAndParts(worksheet.Part);


            //Delete the worksheet part and relation from the package 
            _pck.Workbook.Part.DeleteRelationship(worksheet.RelationshipID);

            //Delete worksheet from the workbook XML
            var sheetsNode = _pck.Workbook.WorkbookXml.SelectSingleNode("//d:workbook/d:sheets", _namespaceManager);
            var sheetNode = sheetsNode?.SelectSingleNode(string.Format("./d:sheet[@sheetId={0}]", worksheet.SheetID), _namespaceManager);
            if (sheetNode != null)
            {
                sheetsNode.RemoveChild(sheetNode);
            }
            _worksheets.Remove(index);
            _pck.Workbook.VbaProject?.Modules.Remove(worksheet.CodeModule);
            ReindexWorksheetDictionary();
            //If the active sheet is deleted, set the first tab as active.
            if (_pck.Workbook.View.ActiveTab >= _pck.Workbook.Worksheets.Count)
            {
                _pck.Workbook.View.ActiveTab = _pck.Workbook.View.ActiveTab-1;
            }
            if (_pck.Workbook.View.ActiveTab == worksheet.SheetID)
            {
                _pck.Workbook.Worksheets[_pck._worksheetAdd].View.TabSelected = true;
            }
            worksheet = null;
        }

        private void DeleteRelationsAndParts(Packaging.ZipPackagePart part)
        {
            var rels = part.GetRelationships().ToList();
            for (var i = 0; i<rels.Count; i++)
            {
                var rel = rels[i];
                if (rel.RelationshipType != ExcelPackage.schemaImage)
                {
                    DeleteRelationsAndParts(_pck.Package.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri)));
                }
                part.DeleteRelationship(rel.Id);
            }
            _pck.Package.DeletePart(part.Uri);
        }

        /// <summary>
        /// Deletes a worksheet from the collection
        /// </summary>
        /// <param name="name">The name of the worksheet in the workbook</param>
        public void Delete(string name)
        {
            var sheet = this[name];
            if (sheet == null)
            {
                throw new ArgumentException(string.Format("Could not find worksheet to delete '{0}'", name));
            }
            Delete(sheet.PositionID);
        }
        /// <summary>
        /// Delete a worksheet from the collection
        /// </summary>
        /// <param name="worksheet">The worksheet to delete</param>
        public void Delete(ExcelWorksheet worksheet)
        {
            if (worksheet.PositionID <= _worksheets.Count && worksheet == _worksheets[worksheet.PositionID])
            {
                Delete(worksheet.PositionID);
            }
            else
            {
                throw new ArgumentException("Worksheet is not in the collection.");
            }
        }
        #endregion
        internal void ReindexWorksheetDictionary()
        {
            var index = _pck._worksheetAdd;
            var worksheets = new Dictionary<int, ExcelWorksheet>();
            foreach (var entry in _worksheets)
            {
                entry.Value.PositionID = index;
                worksheets.Add(index++, entry.Value);
            }
            _worksheets = worksheets;
        }


        /// <summary>
        /// Returns the worksheet at the specified position. 
        /// </summary>
        /// <param name="positionId">The position of the worksheet. Collection is zero-based or one-base depending on the Package.Compatibility.IsWorksheets1Based propery. Default is Zero based</param>
        /// <seealso cref="ExcelPackage.Compatibility"/>
        /// <returns></returns>
        public ExcelWorksheet this[int positionId]
        {
            get
            {
                if (_worksheets.ContainsKey(positionId))
                {
                    return _worksheets[positionId];
                }

                throw new IndexOutOfRangeException("Worksheet position out of range.");
            }
        }

        /// <summary>
        /// Returns the worksheet matching the specified name
        /// </summary>
        /// <param name="name">The name of the worksheet</param>
        /// <returns></returns>
        public ExcelWorksheet this[string name] => GetByName(name);

        /// <summary>
        /// Copies the named worksheet and creates a new worksheet in the same workbook
        /// </summary>
        /// <param name="name">The name of the existing worksheet</param>
        /// <param name="newName">The name of the new worksheet to create</param>
        /// <returns>The new copy added to the end of the worksheets collection</returns>
        public ExcelWorksheet Copy(string name, string newName)
        {
            var copy = this[name];
            if (copy == null)
                throw new ArgumentException(string.Format("Copy worksheet error: Could not find worksheet to copy '{0}'", name));

            var added = Add(newName, copy);
            return added;
        }
        #endregion
        internal ExcelWorksheet GetBySheetID(int localSheetId)
        {
            foreach (var ws in this)
            {
                if (ws.SheetID == localSheetId)
                {
                    return ws;
                }
            }
            return null;
        }
        private ExcelWorksheet GetByName(string name)
        {
            if (string.IsNullOrEmpty(name)) return null;
            ExcelWorksheet xlWorksheet = null;
            foreach (var worksheet in _worksheets.Values)
            {
                if (worksheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    xlWorksheet = worksheet;
            }
            return xlWorksheet;
        }
        #region MoveBefore and MoveAfter Methods
        /// <summary>
        /// Moves the source worksheet to the position before the target worksheet
        /// </summary>
        /// <param name="sourceName">The name of the source worksheet</param>
        /// <param name="targetName">The name of the target worksheet</param>
        public void MoveBefore(string sourceName, string targetName)
        {
            Move(sourceName, targetName, false);
        }

        /// <summary>
        /// Moves the source worksheet to the position before the target worksheet
        /// </summary>
        /// <param name="sourcePositionId">The id of the source worksheet</param>
        /// <param name="targetPositionId">The id of the target worksheet</param>
        public void MoveBefore(int sourcePositionId, int targetPositionId)
        {
            Move(sourcePositionId, targetPositionId, false);
        }

        /// <summary>
        /// Moves the source worksheet to the position after the target worksheet
        /// </summary>
        /// <param name="sourceName">The name of the source worksheet</param>
        /// <param name="targetName">The name of the target worksheet</param>
        public void MoveAfter(string sourceName, string targetName)
        {
            Move(sourceName, targetName, true);
        }

        /// <summary>
        /// Moves the source worksheet to the position after the target worksheet
        /// </summary>
        /// <param name="sourcePositionId">The id of the source worksheet</param>
        /// <param name="targetPositionId">The id of the target worksheet</param>
        public void MoveAfter(int sourcePositionId, int targetPositionId)
        {
            Move(sourcePositionId, targetPositionId, true);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceName"></param>
        public void MoveToStart(string sourceName)
        {
            var sourceSheet = this[sourceName];
            if (sourceSheet == null)
            {
                throw new Exception(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", sourceName));
            }
            Move(sourceSheet.PositionID, _pck._worksheetAdd, false);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourcePositionId"></param>
        public void MoveToStart(int sourcePositionId)
        {
            Move(sourcePositionId, _pck._worksheetAdd, false);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceName"></param>
        public void MoveToEnd(string sourceName)
        {
            var sourceSheet = this[sourceName];
            if (sourceSheet == null)
            {
                throw new Exception(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", sourceName));
            }
            Move(sourceSheet.PositionID, _worksheets.Count + (_pck._worksheetAdd - 1), true);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourcePositionId"></param>
        public void MoveToEnd(int sourcePositionId)
        {
            Move(sourcePositionId, _worksheets.Count+(_pck._worksheetAdd - 1), true);
        }

        private void Move(string sourceName, string targetName, bool placeAfter)
        {
            var sourceSheet = this[sourceName];
            if (sourceSheet == null)
            {
                throw new Exception(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", sourceName));
            }
            var targetSheet = this[targetName];
            if (targetSheet == null)
            {
                throw new Exception(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", targetName));
            }
            Move(sourceSheet.PositionID, targetSheet.PositionID, placeAfter);
        }

        private void Move(int sourcePositionId, int targetPositionId, bool placeAfter)
        {
            // Bugfix: if source and target are the same worksheet the following code will create a duplicate
            //         which will cause a corrupt workbook. /swmal 2014-05-10
            if (sourcePositionId == targetPositionId) return;

            lock (_worksheets)
            {
                var sourceSheet = this[sourcePositionId];
                if (sourceSheet == null)
                {
                    throw new Exception(string.Format("Move worksheet error: Could not find worksheet at position '{0}'", sourcePositionId));
                }
                var targetSheet = this[targetPositionId];
                if (targetSheet == null)
                {
                    throw new Exception(string.Format("Move worksheet error: Could not find worksheet at position '{0}'", targetPositionId));
                }
                if (sourcePositionId == targetPositionId && _worksheets.Count < 2)
                {
                    return;		//--- no reason to attempt to re-arrange a single item with itself
                }

                var index = _pck._worksheetAdd;
                var newOrder = new Dictionary<int, ExcelWorksheet>();
                foreach (var entry in _worksheets)
                {
                    if (entry.Key == targetPositionId)
                    {
                        if (!placeAfter)
                        {
                            sourceSheet.PositionID = index;
                            newOrder.Add(index++, sourceSheet);
                        }

                        entry.Value.PositionID = index;
                        newOrder.Add(index++, entry.Value);

                        if (placeAfter)
                        {
                            sourceSheet.PositionID = index;
                            newOrder.Add(index++, sourceSheet);
                        }
                    }
                    else if (entry.Key == sourcePositionId)
                    {
                        //--- do nothing
                    }
                    else
                    {
                        entry.Value.PositionID = index;
                        newOrder.Add(index++, entry.Value);
                    }
                }
                _worksheets = newOrder;

                MoveSheetXmlNode(sourceSheet, targetSheet, placeAfter);
            }
        }

        private void MoveSheetXmlNode(ExcelWorksheet sourceSheet, ExcelWorksheet targetSheet, bool placeAfter)
        {
            lock (TopNode.OwnerDocument)
            {
                var sourceNode = TopNode.SelectSingleNode(string.Format("d:sheet[@sheetId = '{0}']", sourceSheet.SheetID), _namespaceManager);
                var targetNode = TopNode.SelectSingleNode(string.Format("d:sheet[@sheetId = '{0}']", targetSheet.SheetID), _namespaceManager);
                if (sourceNode == null || targetNode == null)
                {
                    throw new Exception("Source SheetId and Target SheetId must be valid");
                }
                if (placeAfter)
                {
                    TopNode.InsertAfter(sourceNode, targetNode);
                }
                else
                {
                    TopNode.InsertBefore(sourceNode, targetNode);
                }
            }
        }

        #endregion
        public void Dispose()
        {
            if (_worksheets != null)
            {
                foreach (var sheet in _worksheets.Values)
                {
                    ((IDisposable)sheet).Dispose();
                }
                _worksheets = null;
                _pck = null;
            }
        }
    } // end class Worksheets
}
