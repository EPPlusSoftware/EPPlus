/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;
using System.IO;
using System.Linq;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Core.Worksheet;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Core;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.ThreadedComments;
using OfficeOpenXml.Drawing.Slicer;
using System.Text;
using System.Runtime.InteropServices.ComTypes;

namespace OfficeOpenXml
{
    /// <summary>
    /// The collection of worksheets for the workbook
    /// </summary>
    public class ExcelWorksheets : XmlHelper, IEnumerable<ExcelWorksheet>, IDisposable
    {
        #region Private Properties
        private ExcelPackage _pck;
        internal ChangeableDictionary<ExcelWorksheet> _worksheets;
        private XmlNamespaceManager _namespaceManager;
        #endregion
        #region ExcelWorksheets Constructor
        internal ExcelWorksheets(ExcelPackage pck, XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            _pck = pck;
            _namespaceManager = nsm;
            int ix = 0;
            _worksheets = new ChangeableDictionary<ExcelWorksheet>();

            foreach (XmlNode sheetNode in topNode.ChildNodes)
            {
                if (sheetNode.NodeType == XmlNodeType.Element)
                {
                    string name = sheetNode.Attributes["name"].Value;
                    //Get the relationship id
                    string relId = sheetNode.Attributes.GetNamedItem("id", ExcelPackage.schemaRelationships).Value;
                    int sheetID = Convert.ToInt32(sheetNode.Attributes["sheetId"].Value);

                    if (!String.IsNullOrEmpty(relId))
                    {
                        var sheetRelation = pck.Workbook.Part.GetRelationship(relId);
                        Uri uriWorksheet = UriHelper.ResolvePartUri(pck.Workbook.WorkbookUri, sheetRelation.TargetUri);

                        //add the worksheet
                        int positionID = ix + _pck._worksheetAdd;
                        if (sheetRelation.RelationshipType.EndsWith("chartsheet"))
                        {
                            _worksheets.Add(ix, new ExcelChartsheet(_namespaceManager, _pck, relId, uriWorksheet, name, sheetID, positionID, null));
                        }
                        else
                        {
                            _worksheets.Add(ix, new ExcelWorksheet(_namespaceManager, _pck, relId, uriWorksheet, name, sheetID, positionID, null));
                        }
                        ix++;
                    }
                }
            }
        }

        private eWorkSheetHidden TranslateHidden(string value)
        {
            switch (value)
            {
                case "hidden":
                    return eWorkSheetHidden.Hidden;
                case "veryHidden":
                    return eWorkSheetHidden.VeryHidden;
                default:
                    return eWorkSheetHidden.Visible;
            }
        }
        #endregion
        #region ExcelWorksheets Public Properties
        /// <summary>
        /// Returns the number of worksheets in the workbook
        /// </summary>
        public int Count
        {
            get { return (_worksheets.Count); }
        }
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
            return (_worksheets.GetEnumerator());
        }
        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            return (_worksheets.GetEnumerator());
        }

        #endregion
        #region Add Worksheet
        /// <summary>
        /// Adds a new blank worksheet.
        /// </summary>
        /// <param name="Name">The name of the workbook</param>
        public ExcelWorksheet Add(string Name)
        {
            ExcelWorksheet worksheet = AddSheet(Name, false, null);
            return worksheet;
        }
        private ExcelWorksheet AddSheet(string Name, bool isChart, eChartType? chartType, ExcelPivotTable pivotTableSource = null)
        {
            lock (_worksheets)
            {
                Name = ValidateFixSheetName(Name);
                if (GetByName(Name) != null)
                {
                    throw (new InvalidOperationException(ERR_DUP_WORKSHEET + " : " + Name));
                }
                GetSheetURI(ref Name, out int sheetID, out Uri uriWorksheet, isChart);
                Packaging.ZipPackagePart worksheetPart = _pck.ZipPackage.CreatePart(uriWorksheet, isChart ? CHARTSHEET_CONTENTTYPE : WORKSHEET_CONTENTTYPE, _pck.Compression);

                //Create the new, empty worksheet and save it to the package
                StreamWriter streamWorksheet = new StreamWriter(worksheetPart.GetStream(FileMode.Create, FileAccess.Write));
                XmlDocument worksheetXml = CreateNewWorksheet(isChart);
                worksheetXml.Save(streamWorksheet);
                _pck.ZipPackage.Flush();

                string rel = CreateWorkbookRel(Name, sheetID, uriWorksheet, isChart);

                int positionID = _worksheets.Count + _pck._worksheetAdd;
                ExcelWorksheet worksheet;
                if (isChart)
                {
                    worksheet = new ExcelChartsheet(_namespaceManager, _pck, rel, uriWorksheet, Name, sheetID, positionID, eWorkSheetHidden.Visible, (eChartType)chartType, pivotTableSource);
                }
                else
                {
                    worksheet = new ExcelWorksheet(_namespaceManager, _pck, rel, uriWorksheet, Name, sheetID, positionID, eWorkSheetHidden.Visible);
                }

                _worksheets.Add(_worksheets.Count, worksheet);
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
        /// <param name="Copy">The worksheet to be copied</param>
        public ExcelWorksheet Add(string Name, ExcelWorksheet Copy)
        {
            lock (_worksheets)
            {
                int sheetID;
                Uri uriWorksheet;
                if (Copy is ExcelChartsheet)
                {
                    throw (new ArgumentException("Can not copy a chartsheet"));
                }
                if (GetByName(Name) != null)
                {
                    throw (new InvalidOperationException(ERR_DUP_WORKSHEET));
                }

                GetSheetURI(ref Name, out sheetID, out uriWorksheet, false);

                //Create a copy of the worksheet XML
                Packaging.ZipPackagePart worksheetPart = _pck.ZipPackage.CreatePart(uriWorksheet, WORKSHEET_CONTENTTYPE, _pck.Compression);
                StreamWriter streamWorksheet = new StreamWriter(worksheetPart.GetStream(FileMode.Create, FileAccess.Write));

                XmlDocument worksheetXml = new XmlDocument();
                worksheetXml.LoadXml(Copy.WorksheetXml.OuterXml);
                worksheetXml.Save(streamWorksheet);
                _pck.ZipPackage.Flush();

                //Create a relation to the workbook
                string relID = CreateWorkbookRel(Name, sheetID, uriWorksheet, false);

                ExcelWorksheet added = new ExcelWorksheet(_namespaceManager, _pck, relID, uriWorksheet, Name, sheetID, _worksheets.Count + _pck._worksheetAdd, eWorkSheetHidden.Visible);


                //Copy comments
                if (Copy.ThreadedComments.Count > 0)
                {
                    CopyThreadedComments(Copy, added);
                }
                else if (Copy.Comments.Count > 0)
                {
                    CopyComment(Copy, added);
                }
                else if (Copy.VmlDrawingsComments.Count > 0)    //Vml drawings are copied as part of the comments. 
                {
                    CopyVmlDrawing(Copy, added);
                }

                //Copy HeaderFooter
                CopyHeaderFooterPictures(Copy, added);

                //Copy all relationships 
                if (Copy.HasDrawingRelationship)
                {
                    CopySlicers(Copy, added);
                    CopyDrawing(Copy, added);
                }
                if (Copy.Tables.Count > 0)
                {
                    CopyTable(Copy, added);
                }
                if (Copy.PivotTables.Count > 0)
                {
                    CopyPivotTable(Copy, added);
                }
                if (Copy.Names.Count > 0)
                {
                    CopySheetNames(Copy, added);
                }

                //Copy all cells and styles if the worksheet is from another workbook.
                CloneCellsAndStyles(Copy, added);

                //Copy the VBA code
                if (_pck.Workbook.VbaProject != null && Copy.CodeModule != null)
                {
                    var name = _pck.Workbook.VbaProject.GetModuleNameFromWorksheet(added);
                    _pck.Workbook.VbaProject.Modules.Add(new ExcelVBAModule(added.CodeNameChange) { Name = name, Code = Copy.CodeModule.Code, Attributes = _pck.Workbook.VbaProject.GetDocumentAttributes(Name, "0{00020820-0000-0000-C000-000000000046}"), Type = eModuleType.Document, HelpContext = 0 });
                    Copy.CodeModuleName = name;
                }

                _worksheets.Add(_worksheets.Count, added);

                //Remove any relation to printersettings.
                XmlNode pageSetup = added.WorksheetXml.SelectSingleNode("//d:pageSetup", _namespaceManager);
                if (pageSetup != null)
                {
                    XmlAttribute attr = (XmlAttribute)pageSetup.Attributes.GetNamedItem("id", ExcelPackage.schemaRelationships);
                    if (attr != null)
                    {
                        relID = attr.Value;
                        // first delete the attribute from the XML
                        pageSetup.Attributes.Remove(attr);
                    }
                }
                return added;
            }
        }

        private void CopySlicers(ExcelWorksheet copy, ExcelWorksheet added)
        {
            foreach (var source in copy.SlicerXmlSources._list)
            {
                var id = added.SheetId;
                var uri = GetNewUri(added.Part.Package, "/xl/slicers/slicer{0}.xml", ref id);
                var part = added.Part.Package.CreatePart(uri, "application/vnd.ms-excel.slicer+xml", added.Part.Package.Compression);
                var rel = added.Part.CreateRelationship(uri, Packaging.TargetMode.Internal, ExcelPackage.schemaRelationshipsSlicer);
                var xml = new XmlDocument();
                xml.LoadXml(source.XmlDocument.OuterXml);
                var stream = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                xml.Save(stream);

                //Now create the new relationship between the worksheet and the slicer.
                var relNode = (XmlElement)(added.WorksheetXml.DocumentElement.SelectSingleNode($"d:extLst/d:ext/x14:slicerList/x14:slicer[@r:id='{source.Rel.Id}']", NameSpaceManager));
                relNode.Attributes["r:id"].Value = rel.Id;
            }
        }
        /// <summary>
        /// Adds a chartsheet to the workbook.
        /// </summary>
        /// <param name="Name">The name of the worksheet</param>
        /// <param name="chartType">The type of chart</param>
        /// <returns></returns>
        public ExcelChartsheet AddChart(string Name, eChartType chartType)
        {
            if (ExcelChart.IsTypeStock(chartType))
            {
                throw (new InvalidOperationException("Please use method AddStockChart for Stock Charts"));
            }
            return (ExcelChartsheet)AddSheet(Name, true, chartType, null);
        }
        /// <summary>
        /// Adds a chartsheet to the workbook.
        /// </summary>
        /// <param name="Name">The name of the worksheet</param>
        /// <param name="chartType">The type of chart</param>
        /// <param name="pivotTableSource">The pivottable source</param>
        /// <returns></returns>
        public ExcelChartsheet AddChart(string Name, eChartType chartType, ExcelPivotTable pivotTableSource)
        {
            return (ExcelChartsheet)AddSheet(Name, true, chartType, pivotTableSource);
        }
        /// <summary>
        /// Adds a stock chart sheet to the workbook.
        /// </summary>
        /// <param name="Name">The name of the worksheet</param>
        /// <param name="CategorySerie">The category serie. A serie containing dates or names</param>
        /// <param name="HighSerie">The high price serie</param>    
        /// <param name="LowSerie">The low price serie</param>    
        /// <param name="CloseSerie">The close price serie containing</param>    
        /// <param name="OpenSerie">The opening price serie. Supplying this serie will create a StockOHLC or StockVOHLC chart</param>
        /// <param name="VolumeSerie">The volume represented as a column chart. Supplying this serie will create a StockVHLC or StockVOHLC chart</param>
        /// <returns></returns>
        public ExcelChartsheet AddStockChart(string Name, ExcelRangeBase CategorySerie, ExcelRangeBase HighSerie, ExcelRangeBase LowSerie, ExcelRangeBase CloseSerie, ExcelRangeBase OpenSerie = null, ExcelRangeBase VolumeSerie = null)
        {
            var chartType = ExcelStockChart.GetChartType(OpenSerie, VolumeSerie);
            var sheet = (ExcelChartsheet)AddSheet(Name, true, chartType, null);
            var chart = (ExcelStockChart)sheet.Chart;
            ExcelStockChart.SetStockChartSeries(chart, chartType, CategorySerie.FullAddress, HighSerie.FullAddress, LowSerie.FullAddress, CloseSerie.FullAddress, OpenSerie?.FullAddress, VolumeSerie?.FullAddress);
            return sheet;
        }
        private void CopySheetNames(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            foreach (var name in Copy.Names)
            {
                ExcelNamedRange newName;
                if (!name.IsName)
                {
                    if (name.WorkSheetName == Copy.Name)
                    {
                        newName = added.Names.Add(name.Name, added.Cells[name.FirstAddress]);
                    }
                    else
                    {
                        newName = added.Names.Add(name.Name, added.Workbook.Worksheets[name.WorkSheetName].Cells[name.FirstAddress]);
                    }
                }
                else if (!string.IsNullOrEmpty(name.NameFormula))
                {
                    newName = added.Names.AddFormula(name.Name, name.Formula);
                }
                else
                {
                    newName = added.Names.AddValue(name.Name, name.Value);
                }
                newName.NameComment = name.NameComment;
            }
        }

        private void CopyTable(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            string prevName = "";
            //First copy the table XML
            foreach (var tbl in Copy.Tables)
            {
                string xml = tbl.TableXml.OuterXml;
                string name;
                if (prevName == "")
                {
                    name = Copy.Tables.GetNewTableName();
                }
                else
                {
                    int ix = int.Parse(prevName.Substring(5)) + 1;
                    name = string.Format("Table{0}", ix);
                    while (_pck.Workbook.ExistsPivotTableName(name))
                    {
                        name = string.Format("Table{0}", ++ix);
                    }
                }
                //ensure the _nextTableID value has been initialized - Pull request by WillR
                _pck.Workbook.ReadAllTables();

                int Id = _pck.Workbook._nextTableID++;
                prevName = name;
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xml);
                xmlDoc.SelectSingleNode("//d:table/@id", tbl.NameSpaceManager).Value = Id.ToString();
                xmlDoc.SelectSingleNode("//d:table/@name", tbl.NameSpaceManager).Value = name;
                xmlDoc.SelectSingleNode("//d:table/@displayName", tbl.NameSpaceManager).Value = name;
                xml = xmlDoc.OuterXml;

                //var uriTbl = new Uri(string.Format("/xl/tables/table{0}.xml", Id), UriKind.Relative);
                var uriTbl = GetNewUri(_pck.ZipPackage, "/xl/tables/table{0}.xml", ref Id);
                if (_pck.Workbook._nextTableID < Id) _pck.Workbook._nextTableID = Id;

                var part = _pck.ZipPackage.CreatePart(uriTbl, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", _pck.Compression);
                StreamWriter streamTbl = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
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
                    XmlElement elem = added.WorksheetXml.CreateElement("tablePart", ExcelPackage.schemaMain);
                    topNode.AppendChild(elem);
                    elem.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);
                }
                else
                {
                    XmlAttribute relAtt;
                    relAtt = added.WorksheetXml.SelectSingleNode(string.Format("//d:tableParts/d:tablePart/@r:id[.='{0}']", tbl.RelationshipID), tbl.NameSpaceManager) as XmlAttribute;
                    relAtt.Value = rel.Id;
                }

                //Copy table slicers
                foreach (var col in tbl.Columns)
                {
                    if (col.Slicer != null)
                    {
                        var newCol = added.Tables[name].Columns[col.Position];
                        foreach (var d in added.Drawings)
                        {
                            if (d is ExcelTableSlicer slicer)
                            {
                                if (slicer.TableColumn.Name == col.Name && slicer.TableColumn.Table.Id == col.Table.Id)
                                {
                                    slicer.Cache.TableId = newCol.Table.Id;
                                    slicer.TableColumn = newCol;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }
        private void CopyPivotTable(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            _pck.Workbook.ReadAllPivotTables();
            string prevName = "";
            var nameMap = new Dictionary<string, string>();
            foreach (var tbl in Copy.PivotTables)
            {
                string xml = tbl.PivotTableXml.OuterXml;

                string name;
                if (Copy.Workbook == added.Workbook || added.PivotTables._pivotTableNames.ContainsKey(tbl.Name))
                {
                    if (prevName == "")
                    {
                        name = added.PivotTables.GetNewTableName();
                    }
                    else
                    {
                        int ix = int.Parse(prevName.Substring(10)) + 1;
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
                nameMap.Add(tbl.Name, name);
                prevName = name;
                XmlDocument xmlDoc = new XmlDocument();
                //TODO: Fix save pivottable here
                xmlDoc.LoadXml(xml);
                xmlDoc.SelectSingleNode("//d:pivotTableDefinition/@name", tbl.NameSpaceManager).Value = name;
                xml = xmlDoc.OuterXml;

                int Id = _pck.Workbook._nextPivotTableID++;
                var uriTbl = GetNewUri(_pck.ZipPackage, "/xl/pivotTables/pivotTable{0}.xml", ref Id);
                if (_pck.Workbook._nextPivotTableID < Id) _pck.Workbook._nextPivotTableID = Id;
                var partTbl = _pck.ZipPackage.CreatePart(uriTbl, ExcelPackage.schemaPivotTable, _pck.Compression);
                StreamWriter streamTbl = new StreamWriter(partTbl.GetStream(FileMode.Create, FileAccess.Write));
                streamTbl.Write(xml);
                streamTbl.Flush();

                //create the relationship and add the ID to the worksheet xml.
                added.Part.CreateRelationship(UriHelper.ResolvePartUri(added.WorksheetUri, uriTbl), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotTable");
                partTbl.CreateRelationship(tbl.CacheDefinition.CacheDefinitionUri, tbl.CacheDefinition.Relationship.TargetMode, tbl.CacheDefinition.Relationship.RelationshipType);

            }

            added._pivotTables = null;   //Reset collection so it's reloaded when accessing the collection next time.

            //Refresh all items in the copied table.
            foreach (var copiedTbl in added.PivotTables)
            {
                copiedTbl.CacheDefinition._cacheReference._pivotTables.Add(copiedTbl);
                ChangeToWsLocalPivotTable(added, nameMap);
                foreach (var fld in copiedTbl.Fields)
                {
                    fld.Cache.Refresh();
                }
            }
        }

        private void ChangeToWsLocalPivotTable(ExcelWorksheet added, Dictionary<string, string> nameMap)
        {
            foreach(var d in added.Drawings)
            {
                if(d is ExcelPivotTableSlicer s)
                {
                    var list = s.Cache.PivotTables._list;
                    for (int i=0;i<list.Count;i++)
                    {
                        if(nameMap.ContainsKey(list[i].Name))
                        {
                            list[i] = added.PivotTables[nameMap[list[i].Name]];
                        }
                    }
                }
            }
        }

        private void CopyHeaderFooterPictures(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            if (Copy.TopNode != null && Copy.TopNode.SelectSingleNode("d:headerFooter", NameSpaceManager) == null) return;
            //Copy the texts
            if (Copy.HeaderFooter._oddHeader != null) CopyText(Copy.HeaderFooter._oddHeader, added.HeaderFooter.OddHeader);
            if (Copy.HeaderFooter._oddFooter != null) CopyText(Copy.HeaderFooter._oddFooter, added.HeaderFooter.OddFooter);
            if (Copy.HeaderFooter._evenHeader != null) CopyText(Copy.HeaderFooter._evenHeader, added.HeaderFooter.EvenHeader);
            if (Copy.HeaderFooter._evenFooter != null) CopyText(Copy.HeaderFooter._evenFooter, added.HeaderFooter.EvenFooter);
            if (Copy.HeaderFooter._firstHeader != null) CopyText(Copy.HeaderFooter._firstHeader, added.HeaderFooter.FirstHeader);
            if (Copy.HeaderFooter._firstFooter != null) CopyText(Copy.HeaderFooter._firstFooter, added.HeaderFooter.FirstFooter);

            //Copy any images;
            if (Copy.HeaderFooter.Pictures.Count > 0)
            {
                Uri source = Copy.HeaderFooter.Pictures.Uri;
                Uri dest = XmlHelper.GetNewUri(_pck.ZipPackage, @"/xl/drawings/vmlDrawing{0}.vml");
                added.DeleteNode("d:legacyDrawingHF");

                //var part = _pck.Package.CreatePart(dest, "application/vnd.openxmlformats-officedocument.vmlDrawing", _pck.Compression);
                foreach (ExcelVmlDrawingPicture pic in Copy.HeaderFooter.Pictures)
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
            to.LeftAlignedText = from.LeftAlignedText;
            to.CenteredText = from.CenteredText;
            to.RightAlignedText = from.RightAlignedText;
        }
        private void CloneCellsAndStyles(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            bool sameWorkbook = (Copy.Workbook == _pck.Workbook);

            bool doAdjust = _pck.DoAdjustDrawings;
            _pck.DoAdjustDrawings = false;
            //Merged cells
            foreach (var r in Copy.MergedCells)     //Issue #94
            {
                added.MergedCells.Add(new ExcelAddress(r), false);
            }

            //Shared Formulas   
            foreach (int key in Copy._sharedFormulas.Keys)
            {
                added._sharedFormulas.Add(key, Copy._sharedFormulas[key].Clone());
            }

            Dictionary<int, int> styleCashe = new Dictionary<int, int>();
            //Cells
            int row, col;
            var val = new CellStoreEnumerator<ExcelValue>(Copy._values);
            while (val.Next())
            {
                row = val.Row;
                col = val.Column;
                int styleID = 0;
                if (row == 0) //Column
                {
                    var c = Copy.GetValueInner(row, col) as ExcelColumn;
                    if (c != null)
                    {
                        var clone = c.Clone(added, c.ColumnMin);
                        clone.StyleID = c.StyleID;
                        added.SetValueInner(row, col, clone);
                        styleID = c.StyleID;
                    }
                }
                else if (col == 0) //Row
                {
                    var r = Copy.Row(row);
                    if (r != null)
                    {
                        r.Clone(added);
                        styleID = r.StyleID;
                    }

                }
                else
                {
                    styleID = CopyValues(Copy, added, row, col);
                }
                if (!sameWorkbook)
                {
                    if (styleCashe.ContainsKey(styleID))
                    {
                        added.SetStyleInner(row, col, styleCashe[styleID]);
                    }
                    else
                    {
                        var s = added.Workbook.Styles.CloneStyle(Copy.Workbook.Styles, styleID);
                        styleCashe.Add(styleID, s);
                        added.SetStyleInner(row, col, s);
                    }
                }
            }

            //Copy dfx styles used in conditional formatting.
            if (!sameWorkbook)
            {
                CopyDxfStylesConditionalFormatting(Copy, added);
            }

            added._package.DoAdjustDrawings = doAdjust;
        }

        private void CopyDxfStylesConditionalFormatting(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            var dxfStyleCashe = new Dictionary<string, int>();
            for (var i = 0; i < Copy.ConditionalFormatting.Count; i++)
            {
                var cfSource = Copy.ConditionalFormatting[i];
                var dxfElement = ((XmlElement)cfSource.Node);
                var dxfId = dxfElement.GetAttribute("dxfId");
                if (ConvertUtil.TryParseIntString(dxfId, out int dxfIdInt))
                {
                    if (!dxfStyleCashe.ContainsKey(dxfId))
                    {
                        var s = added.Workbook.Styles.CloneDxfStyle(Copy.Workbook.Styles, dxfIdInt);
                        dxfStyleCashe.Add(dxfId, s);
                    }
                }
            }
            var nodes = added.WorksheetXml.SelectNodes("//d:conditionalFormatting/d:cfRule", NameSpaceManager);
            foreach (XmlElement cfRule in nodes)
            {
                var dxfId = cfRule.GetAttribute("dxfId");
                if (dxfStyleCashe.ContainsKey(dxfId))
                {
                    cfRule.SetAttribute("dxfId", dxfStyleCashe[dxfId].ToString());
                }
            }
        }

        private int CopyValues(ExcelWorksheet Copy, ExcelWorksheet added, int row, int col)
        {
            var valueCore = Copy.GetCoreValueInner(row, col);
            added.SetValueStyleIdInner(row, col, valueCore._value, valueCore._styleId);

            byte fl = 0;
            if (Copy._flags.Exists(row, col, ref fl))
            {
                added._flags.SetValue(row, col, fl);
            }

            var v = Copy._formulas.GetValue(row, col);
            if (v != null)
            {
                added.SetFormula(row, col, v);
            }

            var hyperLink = Copy._hyperLinks.GetValue(row, col);
            if (hyperLink != null)
            {
                added._hyperLinks.SetValue(row, col, hyperLink);
            }
            return valueCore._styleId;
        }

        private void CopyThreadedComments(ExcelWorksheet copy, ExcelWorksheet workSheet)
        {
            //Copy the underlaying legacy comments.
            CopyComment(copy, workSheet);

            //First copy the drawing XML
            string xml = copy.ThreadedComments.ThreadedCommentsXml.InnerXml;
            var ix = workSheet.SheetId;
            var tcUri = UriHelper.ResolvePartUri(workSheet.WorksheetUri, GetNewUri(_pck.ZipPackage, "/xl/threadedComments/threadedcomment{0}.xml", ref ix));

            var part = _pck.ZipPackage.CreatePart(tcUri, "application/vnd.ms-excel.threadedcomments+xml", _pck.Compression);

            StreamWriter streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
            streamDrawing.Write(xml);
            streamDrawing.Flush();

            //Add the relationship ID to the worksheet xml.
            workSheet.Part.CreateRelationship(tcUri, Packaging.TargetMode.Internal, ExcelPackage.schemaThreadedComment);

            foreach (var t in workSheet.ThreadedComments._threads)
            {
                for (int i = 0; i < t.Comments.Count; i++)
                {
                    t.Comments[i].Id = ExcelThreadedComment.NewId();
                    if (i == 0)
                    {
                        workSheet.Comments[t.CellAddress].Author = "tc=" + t.Comments[i].Id;
                    }
                    else
                    {
                        t.Comments[i].ParentId = t.Comments[0].Id;
                    }
                }
            }

            if (copy.Workbook != workSheet.Workbook) //Different package. Copy all persons from source package.
            {
                var wbDest = workSheet.Workbook;
                foreach (var p in copy.Workbook.ThreadedCommentPersons)
                {
                    wbDest.ThreadedCommentPersons.Add(p.DisplayName, p.UserId, p.ProviderId, p.Id);
                }
            }
        }
        private void CopyComment(ExcelWorksheet Copy, ExcelWorksheet workSheet)
        {
            //First copy the drawing XML
            string xml = Copy.Comments.CommentXml.InnerXml;
            var uriComment = new Uri(string.Format("/xl/comments{0}.xml", workSheet.SheetId), UriKind.Relative);
            if (_pck.ZipPackage.PartExists(uriComment))
            {
                uriComment = XmlHelper.GetNewUri(_pck.ZipPackage, "/xl/comments{0}.xml");
            }

            var part = _pck.ZipPackage.CreatePart(uriComment, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", _pck.Compression);

            StreamWriter streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
            streamDrawing.Write(xml);
            streamDrawing.Flush();

            //Add the relationship ID to the worksheet xml.
            var commentRelation = workSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uriComment), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/comments");

            xml = Copy.VmlDrawingsComments.VmlDrawingXml.InnerXml;

            var uriVml = new Uri(string.Format("/xl/drawings/vmldrawing{0}.vml", workSheet.SheetId), UriKind.Relative);
            if (_pck.ZipPackage.PartExists(uriVml))
            {
                uriVml = XmlHelper.GetNewUri(_pck.ZipPackage, "/xl/drawings/vmldrawing{0}.vml");
            }

            var vmlPart = _pck.ZipPackage.CreatePart(uriVml, "application/vnd.openxmlformats-officedocument.vmlDrawing", _pck.Compression);
            StreamWriter streamVml = new StreamWriter(vmlPart.GetStream(FileMode.Create, FileAccess.Write));
            streamVml.Write(xml);
            streamVml.Flush();

            var newVmlRel = workSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uriVml), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");

            //Add the relationship ID to the worksheet xml.
            XmlElement e = workSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _namespaceManager) as XmlElement;
            if (e == null)
            {
                workSheet.CreateNode("d:legacyDrawing");
                e = workSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _namespaceManager) as XmlElement;
            }

            e.SetAttribute("id", ExcelPackage.schemaRelationships, newVmlRel.Id);
        }

        internal int? GetFirstVisibleSheetIndex()
        {
            for (int i = 0; i < _worksheets.Count; i++)
            {
                if (_worksheets[i].Hidden == eWorkSheetHidden.Visible)
                {
                    return i;
                }
            }
            throw new InvalidOperationException("The worksheets collection must have at least one visible woreksheet");
        }

        private void CopyDrawing(ExcelWorksheet Copy, ExcelWorksheet workSheet/*, PackageRelationship r*/)
        {
            //First copy the drawing XML                
            string xml = Copy.Drawings.DrawingXml.OuterXml;
            var uriDraw = new Uri(string.Format("/xl/drawings/drawing{0}.xml", workSheet.SheetId), UriKind.Relative);
            var part = _pck.ZipPackage.CreatePart(uriDraw, "application/vnd.openxmlformats-officedocument.drawing+xml", _pck.Compression);
            StreamWriter streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
            streamDrawing.Write(xml);
            streamDrawing.Flush();

            XmlDocument drawXml = new XmlDocument();
            drawXml.LoadXml(xml);
            //Add the relationship ID to the worksheet xml.
            var drawRelation = workSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uriDraw), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/drawing");
            XmlElement e = workSheet.WorksheetXml.SelectSingleNode("//d:drawing", _namespaceManager) as XmlElement;
            e.SetAttribute("id", ExcelPackage.schemaRelationships, drawRelation.Id);
            for (int i = 0; i < Copy.Drawings.Count; i++)
            {
                ExcelDrawing draw = Copy.Drawings[i];
                //draw.AdjustPositionAndSize();       //Adjust position for any change in normal style font/row size etc.
                if (draw is ExcelChart chart)
                {
                    xml = chart.ChartXml.InnerXml;

                    var UriChart = XmlHelper.GetNewUri(_pck.ZipPackage, "/xl/charts/chart{0}.xml");
                    var chartPart = _pck.ZipPackage.CreatePart(UriChart, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", _pck.Compression);
                    StreamWriter streamChart = new StreamWriter(chartPart.GetStream(FileMode.Create, FileAccess.Write));
                    streamChart.Write(xml);
                    streamChart.Flush();
                    //Now create the new relationship to the copied chart xml
                    var prevRelID = draw.TopNode.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart/@r:id", Copy.Drawings.NameSpaceManager).Value;
                    var rel = part.CreateRelationship(UriHelper.GetRelativeUri(uriDraw, UriChart), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/chart");
                    XmlAttribute relAtt = drawXml.SelectSingleNode(string.Format("//c:chart/@r:id[.='{0}']", prevRelID), Copy.Drawings.NameSpaceManager) as XmlAttribute;
                    relAtt.Value = rel.Id;
                }
                else if (draw is ExcelPicture pic)
                {
                    IPictureContainer container = pic;
                    var uri = container.UriPic;
                    var img = PictureStore.ImageToByteArray(pic.Image);
                    var ii = workSheet.Workbook._package.PictureStore.AddImage(img, null, pic.ContentType);

                    var rel = part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, ii.Uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
                    //Fixes problem with invalid image when the same image is used more than once.
                    XmlNode relAtt =
                        drawXml.SelectSingleNode(
                            string.Format(
                                "//xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name[.='{0}']/../../../xdr:blipFill/a:blip/@r:embed",
                                pic.Name), Copy.Drawings.NameSpaceManager);

                    if (relAtt != null)
                    {
                        relAtt.Value = rel.Id;
                    }
                }
                else if (draw is ExcelTableSlicer slicer)
                {
                    var name = _pck.Workbook.GetSlicerName(slicer.Name);
                    var newSlicer = workSheet.Drawings[i] as ExcelTableSlicer;
                    newSlicer.Name = name;
                    newSlicer.SlicerName = name;
                    //The slicer still reference the copied slicers cache. We need to create a new cache for the copied slicer.
                    newSlicer.CreateNewCache();
                }
                else if (draw is ExcelPivotTableSlicer ptSlicer)
                {
                    var name = _pck.Workbook.GetSlicerName(ptSlicer.Name);
                    var newSlicer = workSheet.Drawings[i] as ExcelPivotTableSlicer;
                    newSlicer.Name = name;
                    newSlicer.SlicerName = name;
                    //The slicer still reference the copied slicers cache. We need to create a new cache for the copied slicer.
                    newSlicer.CreateNewCache(ptSlicer.Cache._field);
                }

            }
            //rewrite the drawing xml with the new relID's
            streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
            streamDrawing.Write(drawXml.OuterXml);
            streamDrawing.Flush();

            //Copy the size variables to the copy.
            for (int i = 0; i < Copy.Drawings.Count; i++)
            {
                var draw = Copy.Drawings[i];
                var c = workSheet.Drawings[i];
                if (c != null)
                {
                    c._left = draw._left;
                    c._top = draw._top;
                    c._height = draw._height;
                    c._width = draw._width;
                }
                if (c is ExcelChart chart)
                {
                    for (int j = 0; i < chart.Series.Count; i++)
                    {
                        var s = chart.Series[j];
                        var a = new ExcelAddressBase(s.Series);
                        if (a.WorkSheetName.Equals(Copy.Name))
                        {
                            s.Series = ExcelAddressBase.GetFullAddress(workSheet.Name, a.LocalAddress);
                        }
                        a = new ExcelAddressBase(s.XSeries);
                        if (a.WorkSheetName.Equals(Copy.Name))
                        {
                            s.XSeries = ExcelAddressBase.GetFullAddress(workSheet.Name, a.LocalAddress);
                        }
                    }
                }
            }
        }

        private void CopyVmlDrawing(ExcelWorksheet origSheet, ExcelWorksheet newSheet)
        {
            var xml = origSheet.VmlDrawingsComments.VmlDrawingXml.OuterXml;
            var vmlUri = new Uri(string.Format("/xl/drawings/vmlDrawing{0}.vml", newSheet.SheetId), UriKind.Relative);
            var part = _pck.ZipPackage.CreatePart(vmlUri, "application/vnd.openxmlformats-officedocument.vmlDrawing", _pck.Compression);
            var streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
            streamDrawing.Write(xml);
            streamDrawing.Flush();

            //Add the relationship ID to the worksheet xml.
            var vmlRelation = newSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(newSheet.WorksheetUri, vmlUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");
            var e = newSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _namespaceManager) as XmlElement;
            if (e == null)
            {
                e = newSheet.WorksheetXml.CreateNode(XmlNodeType.Entity, "//d:legacyDrawing", _namespaceManager.LookupNamespace("d")) as XmlElement;
            }
            if (e != null)
            {
                e.SetAttribute("id", ExcelPackage.schemaRelationships, vmlRelation.Id);
            }
        }

        string CreateWorkbookRel(string Name, int sheetID, Uri uriWorksheet, bool isChart)
        {
            //Create the relationship between the workbook and the new worksheet
            var rel = _pck.Workbook.Part.CreateRelationship(UriHelper.GetRelativeUri(_pck.Workbook.WorkbookUri, uriWorksheet), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/" + (isChart ? "chartsheet" : "worksheet"));
            _pck.ZipPackage.Flush();

            //Create the new sheet node
            XmlElement worksheetNode = _pck.Workbook.WorkbookXml.CreateElement("sheet", ExcelPackage.schemaMain);
            worksheetNode.SetAttribute("name", Name);
            worksheetNode.SetAttribute("sheetId", sheetID.ToString());
            worksheetNode.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);

            TopNode.AppendChild(worksheetNode);
            return rel.Id;
        }

        private void GetSheetURI(ref string Name, out int sheetID, out Uri uriWorksheet, bool isChart)
        {
            Name = ValidateFixSheetName(Name);
            sheetID = this.Any() ? this.Max(ws => ws.SheetId) + 1 : 1;
            var uriId = sheetID;


            // get the next available worhsheet uri
            do
            {
                if (isChart)
                {
                    uriWorksheet = new Uri("/xl/chartsheets/chartsheet" + uriId + ".xml", UriKind.Relative);
                }
                else
                {
                    uriWorksheet = new Uri("/xl/worksheets/sheet" + uriId + ".xml", UriKind.Relative);
                }

                uriId++;
            } while (_pck.ZipPackage.PartExists(uriWorksheet));
        }

        internal string ValidateFixSheetName(string Name)
        {
            if (string.IsNullOrEmpty(Name) || Name.Trim() == "")
            {
                throw new ArgumentException("The worksheet can not have an empty name");
            }

            //remove invalid characters
            if (ValidateName(Name))
            {
                if (Name.IndexOf(':') > -1) Name = Name.Replace(":", " ");
                if (Name.IndexOf('/') > -1) Name = Name.Replace("/", " ");
                if (Name.IndexOf('\\') > -1) Name = Name.Replace("\\", " ");
                if (Name.IndexOf('?') > -1) Name = Name.Replace("?", " ");
                if (Name.IndexOf('[') > -1) Name = Name.Replace("[", " ");
                if (Name.IndexOf(']') > -1) Name = Name.Replace("]", " ");
            }

            if (Name.StartsWith("'") || Name.EndsWith("'"))
            {
                throw new ArgumentException("The worksheet name can not start or end with an apostrophe (').", "Name");
            }
            if (Name.Length > 31) Name = Name.Substring(0, 31);   //A sheet can have max 31 char's            
            return Name;
        }
        /// <summary>
        /// Validate the sheetname
        /// </summary>
        /// <param name="Name">The Name</param>
        /// <returns>True if valid</returns>
        private bool ValidateName(string Name)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(Name, @":|\?|/|\\|\[|\]");
        }

        /// <summary>
        /// Creates the XML document representing a new empty worksheet
        /// </summary>
        /// <returns></returns>
        internal XmlDocument CreateNewWorksheet(bool isChart)
        {
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement elemWs = xmlDoc.CreateElement(isChart ? "chartsheet" : "worksheet", ExcelPackage.schemaMain);
            elemWs.SetAttribute("xmlns:r", ExcelPackage.schemaRelationships);
            xmlDoc.AppendChild(elemWs);


            if (isChart)
            {
                XmlElement elemSheetPr = xmlDoc.CreateElement("sheetPr", ExcelPackage.schemaMain);
                elemWs.AppendChild(elemSheetPr);

                XmlElement elemSheetViews = xmlDoc.CreateElement("sheetViews", ExcelPackage.schemaMain);
                elemWs.AppendChild(elemSheetViews);

                XmlElement elemSheetView = xmlDoc.CreateElement("sheetView", ExcelPackage.schemaMain);
                elemSheetView.SetAttribute("workbookViewId", "0");
                elemSheetView.SetAttribute("zoomToFit", "1");

                elemSheetViews.AppendChild(elemSheetView);
            }
            else
            {
                XmlElement elemSheetViews = xmlDoc.CreateElement("sheetViews", ExcelPackage.schemaMain);
                elemWs.AppendChild(elemSheetViews);

                XmlElement elemSheetView = xmlDoc.CreateElement("sheetView", ExcelPackage.schemaMain);
                elemSheetView.SetAttribute("workbookViewId", "0");
                elemSheetViews.AppendChild(elemSheetView);

                XmlElement elemSheetFormatPr = xmlDoc.CreateElement("sheetFormatPr", ExcelPackage.schemaMain);
                elemSheetFormatPr.SetAttribute("defaultRowHeight", "15");
                elemWs.AppendChild(elemSheetFormatPr);

                XmlElement elemSheetData = xmlDoc.CreateElement("sheetData", ExcelPackage.schemaMain);
                elemWs.AppendChild(elemSheetData);
            }
            return xmlDoc;
        }
        #endregion
        #region Delete Worksheet
        /// <summary>
        /// Deletes a worksheet from the collection
        /// </summary>
        /// <param name="Index">The position of the worksheet in the workbook</param>
        public void Delete(int Index)
        {
            /*
            * Hack to prefetch all the drawings,
            * so that all the images are referenced, 
            * to prevent the deletion of the image file, 
            * when referenced more than once
            */
            foreach (var ws in _worksheets)
            {
                var drawings = ws.Drawings;
            }

            ExcelWorksheet worksheet = _worksheets[Index - _pck._worksheetAdd];
            if (worksheet.Drawings.Count > 0)
            {
                worksheet.Drawings.ClearDrawings();
            }

            //Remove all comments
            if (!(worksheet is ExcelChartsheet) && worksheet.Comments.Count > 0)
            {
                worksheet.Comments.Clear();
            }

            while(worksheet.PivotTables.Count>0)
            {
                worksheet.PivotTables.Delete(worksheet.PivotTables[0]);
            }
            //Delete any parts still with relations to the Worksheet.
            DeleteRelationsAndParts(worksheet.Part);


            //Delete the worksheet part and relation from the package 
            _pck.Workbook.Part.DeleteRelationship(worksheet.RelationshipId);

            //Delete worksheet from the workbook XML
            XmlNode sheetsNode = _pck.Workbook.WorkbookXml.SelectSingleNode("//d:workbook/d:sheets", _namespaceManager);
            if (sheetsNode != null)
            {
                XmlNode sheetNode = sheetsNode.SelectSingleNode(string.Format("./d:sheet[@sheetId={0}]", worksheet.SheetId), _namespaceManager);
                if (sheetNode != null)
                {
                    sheetsNode.RemoveChild(sheetNode);
                }
            }
            if (_pck.Workbook.VbaProject != null)
            {
                _pck.Workbook.VbaProject.Modules.Remove(worksheet.CodeModule);
            }

            _worksheets.RemoveAndShift(Index - _pck._worksheetAdd);
            ReindexWorksheetDictionary();
            //If the active sheet is deleted, set the first tab as active.
            if (_pck.Workbook.View.ActiveTab >= _pck.Workbook.Worksheets.Count)
            {
                _pck.Workbook.View.ActiveTab = _pck.Workbook.View.ActiveTab - 1;
            }
            if (_pck.Workbook.View.ActiveTab == worksheet.SheetId)
            {
                _pck.Workbook.Worksheets[_pck._worksheetAdd].View.TabSelected = true;
            }
        }

        private void DeleteRelationsAndParts(Packaging.ZipPackagePart part)
        {
            var rels = part.GetRelationships().ToList();
            for (int i = 0; i < rels.Count; i++)
            {
                var rel = rels[i];
                if (rel.RelationshipType != ExcelPackage.schemaImage && rel.TargetMode == Packaging.TargetMode.Internal)
                {
                    var relUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                    if (_pck.ZipPackage.PartExists(relUri))
                    {
                        DeleteRelationsAndParts(_pck.ZipPackage.GetPart(relUri));
                    }
                }
                part.DeleteRelationship(rel.Id);
            }
            _pck.ZipPackage.DeletePart(part.Uri);
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
            Delete(sheet.PositionId);
        }
        /// <summary>
        /// Delete a worksheet from the collection
        /// </summary>
        /// <param name="Worksheet">The worksheet to delete</param>
        public void Delete(ExcelWorksheet Worksheet)
        {
            var ix = Worksheet.PositionId - _pck._worksheetAdd;
            if (ix < _worksheets.Count && Worksheet == _worksheets[ix])
            {
                Delete(Worksheet.PositionId);
            }
            else
            {
                throw (new ArgumentException("Worksheet is not in the collection."));
            }
        }
        #endregion
        internal void ReindexWorksheetDictionary()
        {
            var index = 0;
            var worksheets = new ChangeableDictionary<ExcelWorksheet>();
            foreach (var entry in _worksheets)
            {
                entry.PositionId = index + _pck._worksheetAdd;
                worksheets.Add(index++, entry);
            }
            _worksheets = worksheets;
        }

#if Core
        /// <summary>
        /// Returns the worksheet at the specified position. 
        /// </summary>
        /// <param name="PositionID">The position of the worksheet. Collection is zero-based or one-base depending on the Package.Compatibility.IsWorksheets1Based propery. Default is Zero based</param>
        /// <seealso cref="ExcelPackage.Compatibility"/>
        /// <returns></returns>
#else
        /// <summary>
        /// Returns the worksheet at the specified position. 
        /// </summary>
        /// <param name="PositionID">The position of the worksheet. Collection is zero-based or one-base depending on the Package.Compatibility.IsWorksheets1Based propery. Default is One based</param>
        /// <seealso cref="ExcelPackage.Compatibility"/>
        /// <returns></returns>
#endif
        public ExcelWorksheet this[int PositionID]
        {
            get
            {
                var ix = PositionID - _pck._worksheetAdd;
                if (_worksheets.ContainsKey(ix))
                {
                    return _worksheets[ix];
                }
                else
                {
                    throw (new IndexOutOfRangeException("Worksheet position out of range."));
                }
            }
        }

        /// <summary>
        /// Returns the worksheet matching the specified name
        /// </summary>
        /// <param name="Name">The name of the worksheet</param>
        /// <returns></returns>
        public ExcelWorksheet this[string Name]
        {
            get
            {
                return GetByName(Name);
            }
        }
        /// <summary>
        /// Copies the named worksheet and creates a new worksheet in the same workbook
        /// </summary>
        /// <param name="Name">The name of the existing worksheet</param>
        /// <param name="NewName">The name of the new worksheet to create</param>
        /// <returns>The new copy added to the end of the worksheets collection</returns>
        public ExcelWorksheet Copy(string Name, string NewName)
        {
            ExcelWorksheet Copy = this[Name];
            if (Copy == null)
                throw new ArgumentException(string.Format("Copy worksheet error: Could not find worksheet to copy '{0}'", Name));

            ExcelWorksheet added = Add(NewName, Copy);
            return added;
        }
        #endregion
        internal ExcelWorksheet GetBySheetID(int localSheetID)
        {
            foreach (ExcelWorksheet ws in this)
            {
                if (ws.SheetId == localSheetID)
                {
                    return ws;
                }
            }
            return null;
        }
        private ExcelWorksheet GetByName(string Name)
        {
            if (string.IsNullOrEmpty(Name)) return null;
            ExcelWorksheet xlWorksheet = null;
            foreach (ExcelWorksheet worksheet in _worksheets)
            {
                if (worksheet.Name.Equals(Name, StringComparison.OrdinalIgnoreCase))
                    xlWorksheet = worksheet;
            }
            return (xlWorksheet);
        }

        /// <summary>
        /// Return a worksheet by its name. Can throw an exception if the worksheet does not exist.
        /// </summary>
        /// <param name="worksheetName">Name of the reqested worksheet</param>
        /// <param name="paramName">Name of the parameter</param>
        /// <param name="throwIfNull">Throws an <see cref="ArgumentNullException"></see> if the worksheet doesn't exist.</param>
        /// <returns></returns>
        private ExcelWorksheet GetWorksheetByName(string worksheetName, string paramName = null, bool throwIfNull = true)
        {
            var worksheet = GetByName(worksheetName);
            if (worksheet == null && throwIfNull)
            {
                throw new ArgumentNullException(paramName ?? "worksheet", $"Could not find worksheet to move sourceName");
            }
            return worksheet;
        }

        //#region Move worksheet functions
        /// <summary>
        /// Moves the source worksheet to the position before the target worksheet
        /// </summary>
        /// <param name="sourceName">The name of the source worksheet</param>
        /// <param name="targetName">The name of the target worksheet</param>
        public void MoveBefore(string sourceName, string targetName)
        {
            MoveSheetXmlNode.RearrangeWorksheets(this, sourceName, targetName, true);
        }

        /// <summary>
        /// Moves the source worksheet to the position before the target worksheet
        /// </summary>
        /// <param name="sourcePositionId">The id of the source worksheet</param>
        /// <param name="targetPositionId">The id of the target worksheet</param>
        public void MoveBefore(int sourcePositionId, int targetPositionId)
        {
            MoveSheetXmlNode.RearrangeWorksheets(this, sourcePositionId, targetPositionId, true);
        }

        /// <summary>
        /// Moves the source worksheet to the position after the target worksheet
        /// </summary>
        /// <param name="sourceName">The name of the source worksheet</param>
        /// <param name="targetName">The name of the target worksheet</param>
        public void MoveAfter(string sourceName, string targetName)
        {
            MoveSheetXmlNode.RearrangeWorksheets(this, sourceName, targetName, false);
        }

        /// <summary>
        /// Moves the source worksheet to the position after the target worksheet
        /// </summary>
        /// <param name="sourcePositionId">The id of the source worksheet</param>
        /// <param name="targetPositionId">The id of the target worksheet</param>
        public void MoveAfter(int sourcePositionId, int targetPositionId)
        {
            MoveSheetXmlNode.RearrangeWorksheets(this, sourcePositionId, targetPositionId, true);
        }

        /// <summary>
        /// Moves the source worksheet to the start of the worksheets collection
        /// </summary>
        /// <param name="sourceName">The name of the source worksheet</param>
        public void MoveToStart(string sourceName)
        {
            Require.Argument(sourceName).IsNotNullOrEmpty("sourceName");
            var worksheet = GetWorksheetByName(sourceName, "sourceName");
            MoveToStart(worksheet.PositionId);
        }
        /// <summary>
        /// Moves the source worksheet to the start of the worksheets collection
        /// </summary>
        /// <param name="sourcePositionId">The position of the source worksheet</param>
        public void MoveToStart(int sourcePositionId)
        {
            MoveSheetXmlNode.RearrangeWorksheets(this, sourcePositionId, _pck._worksheetAdd, true);
        }

        /// <summary>
        /// Moves the source worksheet to the end of the worksheets collection
        /// </summary>
        /// <param name="sourceName">The name of the source worksheet</param>
        public void MoveToEnd(string sourceName)
        {
            Require.Argument(sourceName).IsNotNullOrEmpty("sourceName");
            var worksheet = GetWorksheetByName(sourceName, "sourceName");
            MoveToEnd(worksheet.PositionId);
        }

        /// <summary>
        /// Moves the source worksheet to the end of the worksheets collection
        /// </summary>
        /// <param name="sourcePositionId">The position of the source worksheet</param>
        public void MoveToEnd(int sourcePositionId)
        {
            MoveSheetXmlNode.RearrangeWorksheets(this, sourcePositionId, Count - 1 + _pck._worksheetAdd, false);
        }

        /// <summary>
        /// Dispose the worksheets collection
        /// </summary>
        public void Dispose()
        {
            if (_worksheets != null)
            {
                foreach (var sheet in this._worksheets)
                {
                    ((IDisposable)sheet).Dispose();
                }
                _worksheets = null;
                _pck = null;
            }
        }
    } // end class Worksheets
}
