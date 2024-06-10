/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/14/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using OfficeOpenXml.Packaging;
using System.IO;
using System.Xml;
using OfficeOpenXml.Utils;
using OfficeOpenXml.ThreadedComments;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.VBA;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.ConditionalFormatting;
using System.Xml.Linq;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Linq;

namespace OfficeOpenXml.Core.Worksheet
{
    internal static class WorksheetCopyHelper
    {
        internal static ExcelWorksheet Copy(ExcelWorksheets targetWorksheets, string name, ExcelWorksheet sourceWorksheet)
        {
            int sheetID;
            Uri uriWorksheet;
            if (sourceWorksheet is ExcelChartsheet)
            {
                throw (new ArgumentException("Cannot copy a chartsheet"));
            }
            if (targetWorksheets.GetByName(name) != null)
            {
                throw (new InvalidOperationException(ExcelWorksheets.ERR_DUP_WORKSHEET));
            }

            targetWorksheets.GetSheetURI(ref name, out sheetID, out uriWorksheet, false);
            var pck = targetWorksheets._pck;
            var nsm = targetWorksheets.NameSpaceManager;
            //Create a copy of the worksheet XML
            ZipPackagePart worksheetPart = pck.ZipPackage.CreatePart(uriWorksheet, ExcelWorksheets.WORKSHEET_CONTENTTYPE, pck.Compression);
            StreamWriter streamWorksheet = new StreamWriter(worksheetPart.GetStream(FileMode.Create, FileAccess.Write));

            XmlDocument worksheetXml = new XmlDocument();
            worksheetXml.LoadXml(sourceWorksheet.WorksheetXml.OuterXml);
            worksheetXml.Save(streamWorksheet);
            pck.ZipPackage.Flush();

            //Create a relation to the workbook
            string relID = targetWorksheets.CreateWorkbookRel(name, sheetID, uriWorksheet, false, null);

            ExcelWorksheet targetWorksheet = new ExcelWorksheet(nsm, pck, relID, uriWorksheet, name, sheetID, targetWorksheets.Count + pck._worksheetAdd, eWorkSheetHidden.Visible);

            //Copy comments
            if (sourceWorksheet.ThreadedComments.Count > 0)
            {
                CopyThreadedComments(sourceWorksheet, targetWorksheet);
            }
            else if (sourceWorksheet.Comments.Count > 0)
            {
                CopyComment(sourceWorksheet, targetWorksheet);
            }
            else if (sourceWorksheet.VmlDrawings.Count > 0)    //Vml drawings are copied as part of the comments. 
            {
                CopyVmlDrawing(sourceWorksheet, targetWorksheet);
            }

            //Copy HeaderFooter
            CopyHeaderFooterPictures(sourceWorksheet, targetWorksheet);

            //Copy all relationships 
            if (sourceWorksheet.HasDrawingRelationship)
            {
                CopySlicers(sourceWorksheet, targetWorksheet);
                CopyDrawing(sourceWorksheet, targetWorksheet);
            }
            if (sourceWorksheet.Tables.Count > 0)
            {
                CopyTable(sourceWorksheet, targetWorksheet);
            }

            if (sourceWorksheet.PivotTables.Count > 0)
            {
                CopyPivotTable(sourceWorksheet, targetWorksheet);
            }
            if (sourceWorksheet.Names.Count > 0)
            {
                CopySheetNames(sourceWorksheet, targetWorksheet);
            }
            if(sourceWorksheet.DataValidations.Count > 0) 
            {
                foreach(ExcelDataValidation dv in sourceWorksheet.DataValidations)
                {
                    targetWorksheet.DataValidations.AddCopyOfDataValidation(dv, targetWorksheet);
                }
            }
            if(sourceWorksheet.ConditionalFormatting.Count > 0)
            {
                for (int i = 0; i < sourceWorksheet.ConditionalFormatting.Count; i++)
                {
                    targetWorksheet.ConditionalFormatting.CopyRule(sourceWorksheet.ConditionalFormatting[i]);
                }
            }

            //Copy all cells and styles if the worksheet is from another workbook.
            CloneCellsAndStyles(sourceWorksheet, targetWorksheet);

            //Copy the VBA code

            if (pck.Workbook.VbaProject == null)
            {
                targetWorksheet.CodeModuleName = null;
            }
            else if (sourceWorksheet.CodeModule != null)
            {
                var wsName = pck.Workbook.VbaProject.GetModuleNameFromWorksheet(targetWorksheet);
                pck.Workbook.VbaProject.Modules.Add(
                    new ExcelVBAModule(targetWorksheet.CodeNameChange)
                    {
                        Name = wsName,
                        Code = sourceWorksheet.CodeModule.Code,
                        Attributes = pck.Workbook.VbaProject.GetDocumentAttributes(name, "0{00020820-0000-0000-C000-000000000046}"),
                        Type = eModuleType.Document,
                        HelpContext = 0
                    });

                targetWorksheet.CodeModuleName = wsName;
            }

            SetTableFunction(targetWorksheet);

            targetWorksheets._worksheets.Add(targetWorksheets.Count, targetWorksheet);

            //Remove any relation to printersettings.
            XmlNode pageSetup = targetWorksheet.WorksheetXml.SelectSingleNode("//d:pageSetup", nsm);
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

            return targetWorksheet;
        }

        private static void SetTableFunction(ExcelWorksheet added)
        {
            foreach (var t in added.Tables)
            {
                foreach (var c in t.Columns)
                {
                    if (c.TotalsRowFunction != Table.RowFunctions.None)
                    {
                        t.WorkSheet.SetTableTotalFunction(t, c);
                    }
                }
            }
        }

        private static void CloneCellsAndStyles(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            bool sameWorkbook = (Copy.Workbook == added.Workbook);

            bool doAdjust = added._package.DoAdjustDrawings;
            added._package.DoAdjustDrawings = false;
            //Merged cells
            foreach (var r in Copy.MergedCells)     //Issue #94
            {
                if (r != null)
                {
                    added.MergedCells.Add(new ExcelAddress(r), false);
                }
            }

            //Shared Formulas   
            foreach (int key in Copy._sharedFormulas.Keys)
            {
                var sh = Copy._sharedFormulas[key].Clone();
                sh._ws = added;
				added._sharedFormulas.Add(key, sh);
            }

            Dictionary<int, int> styleCashe = new Dictionary<int, int>();
            bool hasMetadata = Copy._metadataStore.HasValues && sameWorkbook;
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
                    styleID = CopyValues(Copy, added, row, col, hasMetadata);
                }
                if (!sameWorkbook && styleID != 0)
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
                CopyDxfStyles(Copy, added);
            }

            added._package.DoAdjustDrawings = doAdjust;
        }
        internal static void CopyDrawing(ExcelWorksheet source, ExcelWorksheet target)
        {
            var pck = target._package;
            var nsm = target.NameSpaceManager;
            //First copy the drawing XML                
            string xml = source.Drawings.DrawingXml.OuterXml;
            var uriDraw = new Uri(string.Format("/xl/drawings/drawing{0}.xml", target.SheetId), UriKind.Relative);
            var partDraw = pck.ZipPackage.CreatePart(uriDraw, "application/vnd.openxmlformats-officedocument.drawing+xml", pck.Compression);
            StreamWriter streamDrawing = new StreamWriter(partDraw.GetStream(FileMode.Create, FileAccess.Write));
            streamDrawing.Write(xml);
            streamDrawing.Flush();

            XmlDocument drawXml = new XmlDocument();
            drawXml.LoadXml(xml);
            //Add the relationship ID to the worksheet xml.
            var drawRelation = target.Part.CreateRelationship(UriHelper.GetRelativeUri(target.WorksheetUri, uriDraw), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/drawing");
            XmlElement e = target.WorksheetXml.SelectSingleNode("//d:drawing", nsm) as XmlElement;
            e.SetAttribute("id", ExcelPackage.schemaRelationships, drawRelation.Id);
            for (int i = 0; i < source.Drawings.Count; i++)
            {
                var draw = source.Drawings[i];
                CopyDrawingRels(draw, pck, target, partDraw, ref drawXml);
            }

            //rewrite the drawing xml with the new relID's
            streamDrawing = new StreamWriter(partDraw.GetStream(FileMode.Create, FileAccess.Write));
            streamDrawing.Write(drawXml.OuterXml);
            streamDrawing.Flush();

            //Copy the size variables to the copy.
            for (int i = 0; i < source.Drawings.Count; i++)
            {
                var draw = source.Drawings[i];
                var c = target.Drawings[i];
                if (c != null)
                {
                    c._left = draw._left;
                    c._top = draw._top;
                    c._height = draw._height;
                    c._width = draw._width;
                }
                if (c is ExcelChart chart)
                {
                    for (int j = 0; j < chart.Series.Count; j++)
                    {
                        var s = chart.Series[j];
                        if (ExcelAddressBase.IsValidAddress(s.Series))
                        {
                            var a = new ExcelAddressBase(s.Series);
                            if (a.WorkSheetName.Equals(source.Name))
                            {
                                s.Series = ExcelCellBase.GetFullAddress(target.Name, a.LocalAddress);
                            }
                        }
                        if (string.IsNullOrEmpty(s.XSeries) == false && ExcelAddressBase.IsValidAddress(s.XSeries))
                        {
                            var a = new ExcelAddressBase(s.XSeries);
                            if (a.WorkSheetName.Equals(source.Name))
                            {
                                s.XSeries = ExcelCellBase.GetFullAddress(target.Name, a.LocalAddress);
                            }
                        }

                        if (s.HeaderAddress != null && s.HeaderAddress.WorkSheetName.Equals(source.Name))
                        {
                            s.HeaderAddress = new ExcelAddressBase(ExcelCellBase.GetFullAddress(target.Name, s.HeaderAddress.LocalAddress));
                        }

                    }
                }
                else if (c is ExcelTableSlicer slicer)
                {
                    var name = pck.Workbook.GetSlicerName(draw.Name);
                    slicer.Name = name;
                    slicer.SlicerName = name;
                    //The slicer still reference the copied slicers cache. We need to create a new cache for the copied slicer.
                    slicer.CreateNewCache();
                }
                else if (c is ExcelPivotTableSlicer ptSlicer)
                {
                    var name = pck.Workbook.GetSlicerName(draw.Name);
                    ptSlicer.Name = name;
                    ptSlicer.SlicerName = name;
                    //The slicer still reference the copied slicers cache. We need to create a new cache for the copied slicer.
                    ptSlicer.CreateNewCache(((ExcelPivotTableSlicer)draw).Cache._field);
                }

            }
        }

        private static XmlNode GetMatchingNode(XmlDocument drawXml, XmlNode node, XmlNamespaceManager nsm)
        {
            var l = new List<XmlNode>();
            var copyXml = node.OwnerDocument;
            l.Add(node);
            while (node.ParentNode != null)
            {
                node = node.ParentNode;
                l.Insert(0, node);
            }

            XmlNode retNode = drawXml.DocumentElement;
            foreach (XmlElement n in l)
            {
                retNode = retNode.SelectSingleNode(n.Name, nsm);
            }
            return retNode;
        }

        private static void CopyDrawingRels(ExcelDrawing sourceDraw, ExcelPackage pck, ExcelWorksheet target, ZipPackagePart partDraw, ref XmlDocument drawXml)
        {
            //var draw = drawings[i];
            var copy = sourceDraw._drawings.Worksheet;
            var uriDraw = partDraw.Uri;
            if (sourceDraw is ExcelChart chart)
            {
                CopyChartRelations(chart, target, partDraw, drawXml, copy);
            }
            else if (sourceDraw is ExcelPicture pic)
            {
                CopyPicture(target, partDraw, drawXml, copy, pic);
            }
            else if (sourceDraw is ExcelControl ctrl)
            {
                CopyControl(pck, target, ctrl);
            }
            else if (sourceDraw is ExcelShape shp)
            {
                CopyBlipFillDrawing(target, partDraw, drawXml, sourceDraw, shp.Fill, uriDraw);
            }
            else if (sourceDraw is ExcelGroupShape grpDraw)
            {
                for (int j = 0; j < grpDraw.Drawings.Count; j++)
                {
                    CopyDrawingRels(grpDraw.Drawings[j], pck, target, partDraw, ref drawXml);
                }
            }

            if (sourceDraw.HypRel != null)
            {
                ZipPackageRelationship rel;
                if (string.IsNullOrEmpty(sourceDraw.HypRel.Target))
                {
                    rel = partDraw.CreateRelationship(sourceDraw.HypRel.TargetUri, sourceDraw.HypRel.TargetMode, sourceDraw.HypRel.RelationshipType);
                }
                else
                {
                    rel = partDraw.CreateRelationship(sourceDraw.HypRel.Target, sourceDraw.HypRel.TargetMode, sourceDraw.HypRel.RelationshipType);
                }

                XmlNode relAtt =
                    drawXml.SelectSingleNode(
                            $"//{sourceDraw._nvPrPath}[@name='{sourceDraw.Name}']/a:hlinkClick/@r:id", copy.Drawings.NameSpaceManager);

                if (relAtt != null)
                {
                    relAtt.Value = rel.Id;
                }
                
            }
        }

        internal static void CopyControl(ExcelPackage pck, ExcelWorksheet target, ExcelControl ctrl)
        {
            var UriCtrl = XmlHelper.GetNewUri(pck.ZipPackage, "/xl/ctrlProps/ctrlProp{0}.xml");
            var ctrlPart = pck.ZipPackage.CreatePart(UriCtrl, ContentTypes.contentTypeControlProperties, pck.Compression);
            StreamWriter streamChart = new StreamWriter(ctrlPart.GetStream(FileMode.Create, FileAccess.Write));
            streamChart.Write(ctrl.ControlPropertiesXml.OuterXml);
            streamChart.Flush();

            var prevRelID = ctrl._control.RelationshipId;
            var rel = target.Part.CreateRelationship(UriHelper.GetRelativeUri(target.WorksheetUri, UriCtrl), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/ctrlProp");
            var relAtts = target.WorksheetXml.SelectNodes(string.Format("//d:control/@r:id[.='{0}']", prevRelID), target.NameSpaceManager);
            XmlAttribute relAtt = relAtts.Item(relAtts.Count-1) as XmlAttribute; //target.WorksheetXml.SelectSingleNode(string.Format("//d:control/@r:id[.='{0}']", prevRelID), target.NameSpaceManager) as XmlAttribute;
            relAtt.Value = rel.Id;
        }

        internal static void CopyChartRelations(ExcelChart chart, ExcelWorksheet target, ZipPackagePart partDraw, XmlDocument drawXml, ExcelWorksheet source)
        {
            var xml = chart.ChartXml.InnerXml;
            Uri uriChart;
            ZipPackagePart chartPart;
            var targetPck = target._package;
            if (chart._isChartEx)
            {
                uriChart = XmlHelper.GetNewUri(targetPck.ZipPackage, "/xl/charts/chartEx{0}.xml");
                chartPart = targetPck.ZipPackage.CreatePart(uriChart, ContentTypes.contentTypeChartEx, targetPck.Compression);
            }
            else
            {
                uriChart = XmlHelper.GetNewUri(targetPck.ZipPackage, "/xl/charts/chart{0}.xml");
                chartPart = targetPck.ZipPackage.CreatePart(uriChart, ContentTypes.contentTypeChart, targetPck.Compression);
            }
            StreamWriter streamChart = new StreamWriter(chartPart.GetStream(FileMode.Create, FileAccess.Write));
            streamChart.Write(xml);
            streamChart.Flush();
            //Now create the new relationship to the copied chart xml
            XmlNode relNode;
            if (chart._isChartEx)
            {
                relNode = chart.TopNode.SelectSingleNode("mc:AlternateContent/mc:Choice[@Requires='cx1' or @Requires='cx2']/xdr:graphicFrame/a:graphic/a:graphicData/cx:chart/@r:id", source.Drawings.NameSpaceManager);
                string prevRelID = relNode?.Value;
                var rel = partDraw.CreateRelationship(UriHelper.GetRelativeUri(partDraw.Uri, uriChart), Packaging.TargetMode.Internal, ExcelPackage.schemaChartExRelationships);
                XmlAttribute relAtt = drawXml.SelectSingleNode(string.Format("//cx:chart/@r:id[.='{0}']", prevRelID), source.Drawings.NameSpaceManager) as XmlAttribute;
                relAtt.Value = rel.Id;
            }
            else
            {
                relNode = chart.TopNode.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart/@r:id", source.Drawings.NameSpaceManager);
                if(relNode == null)//If null, we check the group shape path instead.
                {
                    relNode = chart.TopNode.SelectSingleNode("a:graphic/a:graphicData/c:chart/@r:id", source.Drawings.NameSpaceManager);
                }
                string prevRelID = relNode?.Value;
                var rel = partDraw.CreateRelationship(UriHelper.GetRelativeUri(partDraw.Uri, uriChart), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/chart");
                XmlAttribute relAtt = drawXml.SelectSingleNode(string.Format("//c:chart/@r:id[.='{0}']", prevRelID), source.Drawings.NameSpaceManager) as XmlAttribute;
                relAtt.Value = rel.Id;
            }
            CopyChartRelations(source, target, chart, chartPart);
        }

        internal static void CopyPicture(ExcelWorksheet added, ZipPackagePart partDraw, XmlDocument drawXml, ExcelWorksheet copy, ExcelPicture pic)
        {
            IPictureContainer container = pic;
            var uri = container.UriPic;
            var ii = added.Workbook._package.PictureStore.AddImage(pic.Image.ImageBytes, null, pic.Image.Type);

            var rel = partDraw.CreateRelationship(UriHelper.GetRelativeUri(added.WorksheetUri, ii.Uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
            //Fixes problem with invalid image when the same image is used more than once.
            XmlNode relAtt =
                drawXml.SelectSingleNode(
                    string.Format(
                        "//xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name[.='{0}']/../../../xdr:blipFill/a:blip/@r:embed",
                        pic.Name), copy.Drawings.NameSpaceManager);

            if (relAtt != null)
            {
                relAtt.Value = rel.Id;
            }
        }

        internal static void CopyChartRelations(ExcelWorksheet copy, ExcelWorksheet added, ExcelChart chart, ZipPackagePart chartPart)
        {
            foreach (var relCopy in chart.Part.GetRelationships())
            {
                var uri = UriHelper.ResolvePartUri(relCopy.SourceUri, relCopy.TargetUri);
                if (relCopy.TargetMode == TargetMode.Internal)
                {
                    if (relCopy.RelationshipType == ExcelPackage.schemaChartStyleRelationships)
                    {
                        uri=XmlHelper.GetNewUri(added._package.ZipPackage, "/xl/charts/style{0}.xml");
                        chartPart.Package.CreatePart(uri, ContentTypes.contentTypeChartStyle, chart.StyleManager.StyleXml.OuterXml);
                    }
                    else if (relCopy.RelationshipType == ExcelPackage.schemaChartColorStyleRelationships)
                    {
                        uri = XmlHelper.GetNewUri(added._package.ZipPackage, "/xl/charts/colors{0}.xml");
                        chartPart.Package.CreatePart(uri, ContentTypes.contentTypeChartColorStyle, chart.StyleManager.ColorsXml.OuterXml);
                    }
                    else if(added.Workbook != copy.Workbook)
                    {
                        if (relCopy.RelationshipType == ExcelPackage.schemaRelationships + "/image")
                        {
                            if (added._package.ZipPackage.PartExists(uri)==false)
                            {
                                var destImgUri=copy._package.ZipPackage.GetPart(uri);
                                var v = added._package.ZipPackage.CreatePart(uri, destImgUri);
                            }
                        }
                    }
                }
                var relAdded = chartPart.CreateRelationship(uri, relCopy.TargetMode, relCopy.RelationshipType);
                relAdded.Id = relCopy.Id;
            }
        }

        internal static void CopyBlipFillDrawing(ExcelWorksheet target, ZipPackagePart targetPart, XmlDocument drawXml, ExcelDrawing draw, ExcelDrawingFill fill, Uri uriDraw)
        {
            if (fill.Style == eFillStyle.BlipFill)
            {
                IPictureContainer container = fill.BlipFill;
                var uri = container.UriPic;
                var img = fill.BlipFill.Image.ImageBytes;
                var ii = target.Workbook._package.PictureStore.AddImage(img, uri, null);

                var rel = targetPart.CreateRelationship(UriHelper.GetRelativeUri(uriDraw, ii.Uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
                //Fixes problem with invalid image when the same image is used more than once.
                XmlNode relAtt =
                    drawXml.SelectSingleNode(
                        string.Format(
                            "//xdr:cNvPr/@name[.='{0}']/../../../xdr:spPr/a:blipFill/a:blip/@r:embed",
                            draw.Name), draw.NameSpaceManager);

                if (relAtt != null)
                {
                    relAtt.Value = rel.Id;
                }
            }
        }

        internal static void CopyVmlDrawing(ExcelWorksheet origSheet, ExcelWorksheet newSheet)
        {
            var xml = origSheet.VmlDrawings.VmlDrawingXml.OuterXml;
            var vmlUri = new Uri(string.Format("/xl/drawings/vmlDrawing{0}.vml", newSheet.SheetId), UriKind.Relative);
            var part = newSheet._package.ZipPackage.CreatePart(vmlUri, "application/vnd.openxmlformats-officedocument.vmlDrawing", newSheet._package.Compression);
            var streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
            streamDrawing.Write(xml);
            streamDrawing.Flush();

            //Add the relationship ID to the worksheet xml.
            var vmlRelation = newSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(newSheet.WorksheetUri, vmlUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");
            var e = newSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", newSheet.NameSpaceManager) as XmlElement;
            if (e == null)
            {
                e = newSheet.WorksheetXml.CreateNode(XmlNodeType.Entity, "//d:legacyDrawing", newSheet.NameSpaceManager.LookupNamespace("d")) as XmlElement;
            }
            if (e != null)
            {
                e.SetAttribute("id", ExcelPackage.schemaRelationships, vmlRelation.Id);
            }
            CopyVmlRelations(origSheet, newSheet);
        }

        private static void CopyComment(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            //First copy the drawing XML
            string xml = Copy.Comments.CommentXml.InnerXml;
            var uriComment = new Uri(string.Format("/xl/comments{0}.xml", added.SheetId), UriKind.Relative);
            if (added._package.ZipPackage.PartExists(uriComment))
            {
                uriComment = XmlHelper.GetNewUri(added._package.ZipPackage, "/xl/comments{0}.xml");
            }

            var part = added._package.ZipPackage.CreatePart(uriComment, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", added._package.Compression);

            StreamWriter streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
            streamDrawing.Write(xml);
            streamDrawing.Flush();

            //Add the relationship ID to the worksheet xml.
            var commentRelation = added.Part.CreateRelationship(UriHelper.GetRelativeUri(added.WorksheetUri, uriComment), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/comments");

            xml = Copy.VmlDrawings.VmlDrawingXml.InnerXml;

            var uriVml = new Uri(string.Format("/xl/drawings/vmldrawing{0}.vml", added.SheetId), UriKind.Relative);
            if (added._package.ZipPackage.PartExists(uriVml))
            {
                uriVml = XmlHelper.GetNewUri(added._package.ZipPackage, "/xl/drawings/vmldrawing{0}.vml");
            }

            var vmlPart = added._package.ZipPackage.CreatePart(uriVml, "application/vnd.openxmlformats-officedocument.vmlDrawing", added._package.Compression);
            StreamWriter streamVml = new StreamWriter(vmlPart.GetStream(FileMode.Create, FileAccess.Write));
            streamVml.Write(xml);
            streamVml.Flush();

            var newVmlRel = added.Part.CreateRelationship(UriHelper.GetRelativeUri(added.WorksheetUri, uriVml), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");

            //Add the relationship ID to the worksheet xml.
            XmlElement e = added.WorksheetXml.SelectSingleNode("//d:legacyDrawing", added.NameSpaceManager) as XmlElement;
            if (e == null)
            {
                added.CreateNode("d:legacyDrawing");
                e = added.WorksheetXml.SelectSingleNode("//d:legacyDrawing", added.NameSpaceManager) as XmlElement;
            }

            e.SetAttribute("id", ExcelPackage.schemaRelationships, newVmlRel.Id);
            added.LoadComments();

            CopyVmlRelations(Copy, added);
        }

        internal static void CopyVmlRelations(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            if (Copy._vmlDrawings.Part == null) return;
            foreach (var r in Copy._vmlDrawings.Part.GetRelationships())
            {
                var newRel = added._vmlDrawings.Part.CreateRelationship(r.TargetUri, r.TargetMode, r.RelationshipType);
                if (newRel.Id != r.Id) //Make sure the id's are the same.
                {
                    newRel.Id = r.Id;
                }
                if (Copy.Workbook != added.Workbook)
                {
                    var uri = UriHelper.ResolvePartUri(r.SourceUri, r.TargetUri);
                    if (!added.Part.Package.PartExists(uri))
                    {                        
                        var sourcePart = Copy._package.ZipPackage.GetPart(uri);
                        added._package.ZipPackage.CreatePart(uri, sourcePart);
                    }
                }
            }
        }

        private static void CopySheetNames(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            var sameWorkbook = Copy.Workbook == added.Workbook;
            foreach (var name in Copy.Names)
            {
                ExcelNamedRange newName;
                if (!name.IsName)
                {
                    if (name.WorkSheetName == Copy.Name)
                    {
                        newName = added.Names.AddName(name.Name, added.Cells[name.LocalAddress]);
                    }
                    else
                    {
                        var wsRef = added.Workbook.Worksheets[name.WorkSheetName];
                        if (wsRef == null)
                        {
                            newName = added.Names.AddFormulaNoValidation(name.Name, "#REF!");
                        }
                        else
                        {
                            newName = added.Names.AddName(name.Name, added.Workbook.Worksheets[name.WorkSheetName].Cells[name.LocalAddress]);
                        }
                    }
                }
                else if (!string.IsNullOrEmpty(name.NameFormula))
                {
                    if(sameWorkbook==false && HasExternalReference(name.Formula))
                    {
                        continue;
					}
                    else
                    {
						newName = added.Names.AddFormulaNoValidation(name.Name, name.Formula);
					}
				}
                else
                {
                    newName = added.Names.AddValue(name.Name, name.Value);
                }
                newName.NameComment = name.NameComment;
            }
        }

		private static bool HasExternalReference(string formula)
		{
			if(formula!=null && formula.IndexOf('[') >= 0)
            {
                    var t=SourceCodeTokenizer.Default.Tokenize(formula);
                return t.Any(x => x.TokenType == TokenType.ExternalReference);
            }
            return false;
		}

		private static void CopyTable(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            string prevName = "";
            //First copy the table XML
            foreach (var tbl in Copy.Tables)
            {
                string xml = tbl.TableXml.OuterXml;
                string name;

                if (Copy.Workbook == added.Workbook || added.Workbook.ExistsTableName(tbl.Name))
                {
                    if (prevName == "")
                    {
                        name = Copy.Tables.GetNewTableName();
                    }
                    else
                    {
                        int ix = int.Parse(prevName.Substring(5)) + 1;
                        name = string.Format("Table{0}", ix);
                        while (added._package.Workbook.ExistsPivotTableName(name))
                        {
                            name = string.Format("Table{0}", ++ix);
                        }
                    }
                }
                else
                {
                    name = tbl.Name;
                }

                //ensure the _nextTableID value has been initialized - Pull request by WillR
                added.Workbook.ReadAllTables();

                int Id = added.Workbook._nextTableID++;
                prevName = name;
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xml);
                xmlDoc.SelectSingleNode("//d:table/@id", tbl.NameSpaceManager).Value = Id.ToString();
                xmlDoc.SelectSingleNode("//d:table/@name", tbl.NameSpaceManager).Value = name;
                xmlDoc.SelectSingleNode("//d:table/@displayName", tbl.NameSpaceManager).Value = name;
                xml = xmlDoc.OuterXml;

                //var uriTbl = new Uri(string.Format("/xl/tables/table{0}.xml", Id), UriKind.Relative);
                var uriTbl = XmlHelper.GetNewUri(added._package.ZipPackage, "/xl/tables/table{0}.xml", ref Id);
                if (added.Workbook._nextTableID < Id) added.Workbook._nextTableID = Id;

                var part = added._package.ZipPackage.CreatePart(uriTbl, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", added._package.Compression);
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
        private static void CopyPivotTable(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            Copy._package.Workbook.ReadAllPivotTables();
            string prevName = "";
            var worksheetMap = new Dictionary<string, string>();
            var nameMap = new Dictionary<string, string>();
            var wbAdded = added.Workbook;
            var isPackageInternal = Copy.Workbook == wbAdded;
            foreach (var tbl in Copy.PivotTables)
            {
                string xml = tbl.PivotTableXml.OuterXml;
                string name;
                if (isPackageInternal || added.PivotTables._pivotTableNames.ContainsKey(tbl.Name))
                {
                    if (prevName == "")
                    {
                        name = added.PivotTables.GetNewTableName();
                    }
                    else
                    {
                        int ix = int.Parse(prevName.Substring(10)) + 1;
                        name = string.Format("PivotTable{0}", ix);
                        while (added.Workbook.ExistsPivotTableName(name))
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

                int Id = added.Workbook._nextPivotTableID++;
                var uriTbl = XmlHelper.GetNewUri(added._package.ZipPackage, "/xl/pivotTables/pivotTable{0}.xml", ref Id);
                if (added.Workbook._nextPivotTableID < Id) added.Workbook._nextPivotTableID = Id;

                xml = xmlDoc.OuterXml;

                var partTbl = added._package.ZipPackage.CreatePart(uriTbl, ContentTypes.contentTypePivotTable, added._package.Compression);
                StreamWriter streamTbl = new StreamWriter(partTbl.GetStream(FileMode.Create, FileAccess.Write));
                streamTbl.Write(xml);
                streamTbl.Flush();

                //create the relationship and add the ID to the worksheet xml.
                added.Part.CreateRelationship(UriHelper.ResolvePartUri(added.WorksheetUri, uriTbl), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotTable");
                if (isPackageInternal)
                {
                    partTbl.CreateRelationship(tbl.CacheDefinition.CacheDefinitionUri, tbl.CacheDefinition.Relationship.TargetMode, tbl.CacheDefinition.Relationship.RelationshipType);
                }
                else
                {
                    CreateCacheInNewPackage(added, tbl, partTbl);
                }

            }

            added._pivotTables = null;   //Reset collection so it's reloaded when accessing the collection next time.

            //Refresh all items in the copied table.
            foreach (var copiedTbl in added.PivotTables)
            {
                if (!copiedTbl.CacheDefinition._cacheReference._pivotTables.Contains(copiedTbl))
                {
                    copiedTbl.CacheDefinition._cacheReference._pivotTables.Add(copiedTbl);
                }
                
                if(isPackageInternal==false)
                {
                    copiedTbl.CacheId = copiedTbl.CacheDefinition._cacheReference.CacheId;
                }

                if (copiedTbl.CacheDefinition.IsExternalReferernce) continue;
                
                ChangeToWsLocalPivotTable(added, nameMap);
                foreach (var fld in copiedTbl.Fields)
                {
                    fld.Cache.Refresh();
                }
            }
            //Can't have a cell selected when "group editing" avoids pop-up by not selecting sheet.
            added.View.SetTabSelected(false);
        }

        private static void CreateCacheInNewPackage(ExcelWorksheet added, ExcelPivotTable tbl, ZipPackagePart partTbl)
        {
            var wbAdded = added.Workbook;
            PivotTableCacheInternal newCache;
            var cacheAddress = tbl.CacheDefinition._cacheReference.GetSourceAddress();
            if (wbAdded._pivotTableCaches.TryGetValue(cacheAddress, out ExcelWorkbook.PivotTableCacheRangeInfo rangeInfo))
            {
                newCache = rangeInfo.PivotCaches[0];
                partTbl.CreateRelationship(newCache.CacheDefinitionUri, tbl.CacheDefinition.Relationship.TargetMode, tbl.CacheDefinition.Relationship.RelationshipType);                
            }
            else
            {
                rangeInfo = new ExcelWorkbook.PivotTableCacheRangeInfo();
                string xmlCache = tbl.CacheDefinition.CacheDefinitionXml.OuterXml;
                var cacheId = wbAdded._nextPivotCacheId;
                var uriCache = XmlHelper.GetNewUri(added._package.ZipPackage, "/xl/pivotCache/pivotCacheDefinition{0}.xml", ref cacheId);
                if (wbAdded._nextPivotCacheId < cacheId) wbAdded._nextPivotCacheId = cacheId;

                var partCache = added._package.ZipPackage.CreatePart(uriCache, ContentTypes.contentTypePivotCacheDefinition, added._package.Compression);
                StreamWriter streamCache = new StreamWriter(partCache.GetStream(FileMode.Create, FileAccess.Write));
                streamCache.Write(xmlCache);
                streamCache.Flush();
                partTbl.CreateRelationship(uriCache, tbl.CacheDefinition.Relationship.TargetMode, tbl.CacheDefinition.Relationship.RelationshipType);

                newCache = new PivotTableCacheInternal(wbAdded, uriCache, cacheId);
                rangeInfo.PivotCaches = new List<PivotTableCacheInternal>();

                if (tbl.CacheDefinition.SourceRange != null)
                {
                    if (tbl.CacheDefinition.SourceRange.Worksheet != null && tbl.CacheDefinition.SourceRange.Worksheet.Name == tbl.WorkSheet.Name)
                    {
                        rangeInfo.Address = ExcelCellBase.GetQuotedWorksheetName(added.Name) + "!" + tbl.CacheDefinition.SourceRange.LocalAddress;
                        newCache.SetXmlNodeString(PivotTableCacheInternal._sourceWorksheetPath, added.Name);
                    }
                    else
                    {
                        rangeInfo.Address = cacheAddress;
                    }
                }
                if(tbl.CacheDefinition.SourceExternalReference!=null)
                {
                    var rel = tbl.CacheDefinition._cacheReference.Part.GetRelationship(tbl.CacheDefinition._cacheReference.SourceRId);
                    var rId = partCache.CreateRelationship(rel.TargetUri, rel.TargetMode, rel.RelationshipType);
                    newCache.SourceRId = rId.Id;
                }
                added.Workbook.AddPivotTableCache(newCache, true);

                newCache.AddRecordsXml();
            }
        }

        private static void ChangeToWsLocalPivotTable(ExcelWorksheet added, Dictionary<string, string> nameMap)
        {
            foreach (var d in added.Drawings)
            {
                if (d is ExcelPivotTableSlicer s)
                {
                    var list = s.Cache.PivotTables._list;
                    for (int i = 0; i < list.Count; i++)
                    {
                        if (nameMap.ContainsKey(list[i].Name))
                        {
                            list[i] = added.PivotTables[nameMap[list[i].Name]];
                        }
                    }
                }
            }
        }
        private static void CopyDxfStyles(ExcelWorksheet copy, ExcelWorksheet added)
        {
            //DxfStyleHandler.UpdateDxfXml(copy.Workbook);

            var dxfStyleCashe = new Dictionary<int, int>();
            CopyDxfStylesTables(copy, added);
            CopyDxfStylesPivotTables(copy, added, dxfStyleCashe);
            CopyDxfStylesConditionalFormatting(copy, added, dxfStyleCashe);
        }
        private static void CopyDxfStylesTables(ExcelWorksheet copy, ExcelWorksheet added)
        {
            //Table formats
            for(int i=0;i<copy.Tables.Count; i++)
            {
                var tblFrom = copy.Tables[i];
                var tblTo = added.Tables[tblFrom.Name]; //Use Name, as id can differ if the worksheets are in different workbooks.
                if (tblFrom.HeaderRowStyle.HasValue) tblTo.HeaderRowStyle = (ExcelDxfStyle)tblFrom.HeaderRowStyle.Clone();
                if (tblFrom.HeaderRowBorderStyle.HasValue) tblTo.HeaderRowBorderStyle = (ExcelDxfBorderBase)tblFrom.HeaderRowBorderStyle.Clone();
                if (tblFrom.DataStyle.HasValue) tblTo.DataStyle = (ExcelDxfStyle)tblFrom.DataStyle.Clone();
                if (tblFrom.TableBorderStyle.HasValue) tblTo.TableBorderStyle = (ExcelDxfBorderBase)tblFrom.TableBorderStyle.Clone();
                if (tblFrom.TotalsRowStyle.HasValue) tblTo.TotalsRowStyle = (ExcelDxfStyle)tblFrom.TotalsRowStyle.Clone();
                for (int c=0;c < tblFrom.Columns.Count;c++)
                {
                    var colFrom = tblFrom.Columns[c];
                    var colTo = tblTo.Columns[c];
                    if (colFrom.HeaderRowStyle.HasValue) colTo.HeaderRowStyle = (ExcelDxfStyle)colFrom.HeaderRowStyle.Clone();
                    if (colFrom.DataStyle.HasValue) colTo.DataStyle = (ExcelDxfStyle)colFrom.DataStyle.Clone();
                    if (colFrom.TotalsRowStyle.HasValue) colTo.TotalsRowStyle = (ExcelDxfStyle)colFrom.TotalsRowStyle.Clone();
                }
            }
        }
        private static void CopyDxfStylesPivotTables(ExcelWorksheet copy, ExcelWorksheet added, Dictionary<int, int> dxfStyleCache)
        {
            //Table formats
            foreach (var pt in copy.PivotTables)
            {
                var ix = 0;
                var newPt = added.PivotTables[pt.Name];
                foreach (var a in pt.Styles._list)
                {
                    var addedStyle = newPt.Styles[ix++];
                    addedStyle.DxfId = int.MinValue;                    
                    addedStyle.Style = (ExcelDxfStyle)a.Style.Clone();
                }                
            }
        }
        private static void CopyDxfStylesConditionalFormatting(ExcelWorksheet copy, ExcelWorksheet added, Dictionary<int, int> dxfStyleCache)
        {
            //Conditional Formatting
            for (var i = 0; i < copy.ConditionalFormatting.Count; i++)
            {
                var cfSource = copy.ConditionalFormatting[i];
                var dxfId = cfSource.DxfId;
                if (dxfId != -1)
                {
                    AppendDxf(copy.Workbook.Styles, added.Workbook.Styles, dxfStyleCache, dxfId);
                    added.ConditionalFormatting[i].DxfId = dxfStyleCache[dxfId];
                }
            }
        }

        private static void AppendDxf(ExcelStyles stylesFrom, ExcelStyles stylesTo, Dictionary<int, int> dxfStyleCache, int dxfId)
        {
            if (dxfId < 0) return;
            if (!dxfStyleCache.ContainsKey(dxfId))
            {
                var s = DxfStyleHandler.CloneDxfStyle(stylesFrom, stylesTo, dxfId, ExcelStyles.DxfsPath);
                dxfStyleCache.Add(dxfId, s);
            }
        }

        private static int CopyValues(ExcelWorksheet Copy, ExcelWorksheet added, int row, int col, bool hasMetadata)
        {
            var valueCore = Copy.GetCoreValueInner(row, col);
            added.SetValueStyleIdInner(row, col, valueCore._value, valueCore._styleId);

            byte fl = 0;
            if (Copy._flags.Exists(row, col, ref fl))
            {
                added._flags.SetValue(row, col, fl);
            }
            if (hasMetadata)
            {
                ExcelWorksheet.MetaDataReference md = new ExcelWorksheet.MetaDataReference();
                if (Copy._metadataStore.Exists(row, col, ref md))
                {
                    added._metadataStore.SetValue(row, col, md);
                }
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

        private static void CopyThreadedComments(ExcelWorksheet copy, ExcelWorksheet added)
        {
            //Copy the underlaying legacy comments.
            CopyComment(copy, added);

            //First copy the drawing XML
            string xml = copy.ThreadedComments.ThreadedCommentsXml.InnerXml;
            var ix = added.SheetId;
            var tcUri = UriHelper.ResolvePartUri(added.WorksheetUri, XmlHelper.GetNewUri(added._package.ZipPackage, "/xl/threadedComments/threadedComment{0}.xml", ref ix));

            var part = added._package.ZipPackage.CreatePart(tcUri, "application/vnd.ms-excel.threadedcomments+xml", added._package.Compression);

            StreamWriter streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
            streamDrawing.Write(xml);
            streamDrawing.Flush();

            //Add the relationship ID to the worksheet xml.
            added.Part.CreateRelationship(tcUri, Packaging.TargetMode.Internal, ExcelPackage.schemaThreadedComment);

            added.LoadThreadedComments();
            foreach (var t in added.ThreadedComments._threads)
            {
                for (int i = 0; i < t.Comments.Count; i++)
                {
                    t.Comments[i].Id = ExcelThreadedComment.NewId();
                    if (i == 0)
                    {
                        added.Comments[t.CellAddress].Author = "tc=" + t.Comments[i].Id;
                    }
                    else
                    {
                        t.Comments[i].ParentId = t.Comments[0].Id;
                    }
                }
            }

            if (copy.Workbook != added.Workbook) //Different package. Copy all persons from source package.
            {
                var wbDest = added.Workbook;
                foreach (var p in copy.Workbook.ThreadedCommentPersons)
                {
                    wbDest.ThreadedCommentPersons.Add(p.DisplayName, p.UserId, p.ProviderId, p.Id);
                }
            }
        }
        private static void CopyHeaderFooterPictures(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            if (Copy.TopNode != null && Copy.GetNode("d:headerFooter") == null) return;
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
                Uri dest = XmlHelper.GetNewUri(added._package.ZipPackage, @"/xl/drawings/vmlDrawing{0}.vml");
                added.DeleteNode("d:legacyDrawingHF");

                foreach (ExcelVmlDrawingPicture pic in Copy.HeaderFooter.Pictures)
                {
                    ExcelVmlDrawingPicture item;
                    if (Copy._package != added._package)
                    {
                        var ii = added.Workbook._package.PictureStore.AddImage(pic.Image.ImageBytes, null, pic.Image.Type);
                        item = added.HeaderFooter.Pictures.Add(pic.Id, ii.Uri, pic.Title, pic.Width, pic.Height);
                    }
                    else
                    {
                        item = added.HeaderFooter.Pictures.Add(pic.Id, pic.ImageUri, pic.Title, pic.Width, pic.Height);
                    }
                    foreach (XmlAttribute att in pic.TopNode.Attributes)
                    {
                        (item.TopNode as XmlElement).SetAttribute(att.Name, att.Value);
                    }
                    item.TopNode.InnerXml = pic.TopNode.InnerXml;
                }
            }
        }
        private static void CopyText(ExcelHeaderFooterText from, ExcelHeaderFooterText to)
        {
            to.LeftAlignedText = from.LeftAlignedText;
            to.CenteredText = from.CenteredText;
            to.RightAlignedText = from.RightAlignedText;
        }

        private static void CopySlicers(ExcelWorksheet source, ExcelWorksheet target)
        {
            foreach (var slicer in source.SlicerXmlSources._list)
            {
                var id = target.SheetId;
                var uri = XmlHelper.GetNewUri(target.Part.Package, "/xl/slicers/slicer{0}.xml", ref id);
                var part = target.Part.Package.CreatePart(uri, "application/vnd.ms-excel.slicer+xml", target.Part.Package.Compression);
                var rel = target.Part.CreateRelationship(uri, Packaging.TargetMode.Internal, ExcelPackage.schemaRelationshipsSlicer);
                var xml = new XmlDocument();
                xml.LoadXml(slicer.XmlDocument.OuterXml);
                var stream = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                xml.Save(stream);

                //Now create the new relationship between the worksheet and the slicer.
                var relNode = (XmlElement)(target.WorksheetXml.DocumentElement.SelectSingleNode($"d:extLst/d:ext/x14:slicerList/x14:slicer[@r:id='{slicer.Rel.Id}']", target.NameSpaceManager));
                relNode.Attributes["r:id"].Value = rel.Id;
            }
        }
    }
}
