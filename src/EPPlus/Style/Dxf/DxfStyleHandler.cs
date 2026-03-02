/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/29/2021         EPPlus Software AB       EPPlus 5.6
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf
{
    internal static class DxfStyleHandler
    {        
        internal static void Load(ExcelWorkbook wb, ExcelStyles styles, ExcelStyleCollection<ExcelDxfStyleBase> dxfs, string path)
        {
            //dxfsPath
            XmlNode dxfsNode = styles.GetNode(path);
            if (dxfsNode != null)
            {
                foreach (XmlNode x in dxfsNode)
                {
                    var item = new ExcelDxfStyle(styles.NameSpaceManager, x, styles, null);
                    dxfs.Add(item.Id, item);
                }
            }
        }
        internal static int CloneDxfStyle(ExcelStyles stylesFrom, ExcelStyles stylesTo, int styleId, string path)
        {
            var copy = stylesFrom.Dxfs[styleId];
            var ix = stylesTo.Dxfs.FindIndexById(copy.Id);
            if (ix < 0)
            {
                var parent = stylesTo.GetNode(path);
                var node = stylesTo.TopNode.OwnerDocument.CreateElement("d:dxf", ExcelPackage.schemaMain);
                parent.AppendChild(node);
                node.InnerXml = copy._helper.TopNode.InnerXml;                
                var dxf = new ExcelDxfStyle(stylesTo.NameSpaceManager, node, stylesTo, null);
                stylesTo.Dxfs.Add(copy.Id, dxf);
                return stylesTo.Dxfs.Count - 1;
            }
            else
            {
                return ix;
            }
        }

        internal static void UpdateDxfXml(ExcelWorkbook wb)
        {
            //Set dxf styling for conditional Formatting
            XmlNode dxfsNode = wb.Styles.TopNode.SelectSingleNode(ExcelStyles.DxfsPath, wb.NameSpaceManager);
            
            UpdateTableStyles(wb, wb.Styles, dxfsNode);
            UpdateDxfXmlWorksheet(wb, wb.Styles, dxfsNode);            
            if (dxfsNode != null) (dxfsNode as XmlElement).SetAttribute("count", wb.Styles.Dxfs.Count.ToString());

            UpdateSlicerStyles(wb, wb.Styles, dxfsNode);
        }

        private static void UpdateTableStyles(ExcelWorkbook wb, ExcelStyles styles, XmlNode dxfsNode)
        {
            foreach (var ts in styles.TableStyles)
            {
                foreach(var element in ts._dic.Values)
                {
                    AddDxfNode(styles.Dxfs, dxfsNode, element.Style);
                    if(element.Style.DxfId>=0)
                    {
                        element.CreateNode();
                    }
                }
            }
        }
        private static void UpdateSlicerStyles(ExcelWorkbook wb, ExcelStyles styles, XmlNode dxfsNode)
        {
            if (styles.SlicerStyles.Count > 0)
            {
                XmlNode extNode = wb.Styles.GetOrCreateExtLstSubNode(ExtLstUris.SlicerStylesDxfCollectionUri, "x14");
                var helper = XmlHelperFactory.Create(styles.NameSpaceManager, extNode);
                var dxfsSlicerNode = helper.CreateNode("x14:dxfs");
                foreach (var ts in styles.SlicerStyles)
                {
                    foreach (var element in ts._dicTable.Values)
                    {
                        AddDxfNode(styles.Dxfs, dxfsNode, element.Style);
                        if (element.Style.DxfId >= 0)
                        {
                            element.CreateNode();
                        }
                    }
                    foreach (var element in ts._dicSlicer.Values)
                    {
                        AddDxfNode(styles.DxfsSlicers, dxfsSlicerNode, element.Style);
                        if (element.Style.DxfId >= 0)
                        {
                            element.CreateNode();
                        }
                    }
                }
            }
        }

        private static void UpdateDxfXmlWorksheet(ExcelWorkbook wb, ExcelStyles styles, XmlNode dxfsNode)
        {
            foreach (var ws in wb.Worksheets)
            {
                if (ws is ExcelChartsheet) continue;
                UpdateConditionalFormatting(ws, styles.Dxfs, dxfsNode);
                UpdateDxfXmlTables(styles, dxfsNode, ws);
                UpdateDxfXmlPivotTables(styles, dxfsNode, ws);
            }
        }
        private static void UpdateDxfXmlTables(ExcelStyles styles, XmlNode dxfsNode, ExcelWorksheet ws)
        {
            foreach (var tbl in ws.Tables)
            {
                tbl.HeaderRowDxfId = AddDxfNode(styles.Dxfs, dxfsNode, tbl.HeaderRowStyle);
                tbl.DataDxfId = AddDxfNode(styles.Dxfs, dxfsNode, tbl.DataStyle);
                tbl.TotalsRowDxfId = AddDxfNode(styles.Dxfs, dxfsNode, tbl.TotalsRowStyle);
                
                tbl.HeaderRowBorderDxfId = AddDxfBorderNode(styles, dxfsNode, tbl.HeaderRowBorderStyle);
                tbl.TableBorderDxfId = AddDxfBorderNode(styles, dxfsNode, tbl.TableBorderStyle);

                foreach (var column in tbl.Columns)
                {
                    column.HeaderRowDxfId = AddDxfNode(styles.Dxfs, dxfsNode, column.HeaderRowStyle);
                    column.DataDxfId = AddDxfNode(styles.Dxfs, dxfsNode, column.DataStyle);
                    column.TotalsRowDxfId = AddDxfNode(styles.Dxfs, dxfsNode, column.TotalsRowStyle);
                }
            }
        }

        private static void UpdateDxfXmlPivotTables(ExcelStyles styles, XmlNode dxfsNode, ExcelWorksheet ws)
        {
            if (ws.HasLoadedPivotTables == false) return;
            foreach (var pt in ws.PivotTables)
            {
                for(int i= 0; i< pt.Styles.Count;i++)
                {
                    var pas = pt.Styles[i];
                    if (pas.Style.HasValue)
                    {
                        pas.DxfId = AddDxfNode(styles.Dxfs, dxfsNode, pas.Style);
                    }
                    else
                    {
                        pt.Styles._list.Remove(pas);   //No dxf style set. We remove the area.
                        i--;
                    }
                }
            }
        }

        private static int? AddDxfBorderNode(ExcelStyles styles, XmlNode dxfsNode, ExcelDxfBorderBase borderStyle)
        {
            if(borderStyle.HasValue)
            {
                var ix = styles.Dxfs.FindIndexById(borderStyle.Id);
                if (ix < 0)
                {
                    var elem = dxfsNode.OwnerDocument.CreateElement("dxf", ExcelPackage.schemaMain);
                    borderStyle.CreateNodes(new XmlHelperInstance(styles.NameSpaceManager, elem), "d:border");
                    dxfsNode.AppendChild(elem);
                    var dxfId = styles.Dxfs.Count;
                    styles.Dxfs.Add(borderStyle.Id, new ExcelDxfTableStyle(styles.NameSpaceManager, elem, styles) { Border = borderStyle });
                    return styles.Dxfs.Count - 1;
                }
                else
                {
                    return ix;
                }
            }
            return null;
        }
        private static int? AddDxfNode(ExcelStyleCollection<ExcelDxfStyleBase> dxfs, XmlNode dxfsNode, ExcelDxfStyleBase dxfStyle)
        {
            if (dxfStyle.HasValue)
            {
                var ix = dxfs.FindIndexById(dxfStyle.Id);
                if (ix < 0)
                {
                    dxfStyle.DxfId = dxfs.Count;
                    dxfs.Add(dxfStyle.Id, dxfStyle);
                    var elem = dxfsNode.OwnerDocument.CreateElement("dxf", ExcelPackage.schemaMain);
                    dxfStyle.CreateNodes(new XmlHelperInstance(dxfStyle._helper.NameSpaceManager, elem), "");
                    dxfsNode.AppendChild(elem);
                }
                else
                {
                    dxfStyle.DxfId = ix;
                }
                return dxfStyle.DxfId;
            }
            return null;
        }

        private static void UpdateConditionalFormatting(ExcelWorksheet ws, ExcelStyleCollection<ExcelDxfStyleBase> dxfs, XmlNode dxfsNode)
        {
            foreach (var cf in ws.ConditionalFormatting)
            {
                if (cf.Style.HasValue)
                {
                    int ix = dxfs.FindIndexById(cf.Style.Id);
                    if (ix < 0)
                    {
                        ((ExcelConditionalFormattingRule)cf).DxfId = dxfs.Count;
                        dxfs.Add(cf.Style.Id, cf.Style);
                        var elem = dxfsNode.OwnerDocument.CreateElement("dxf", ExcelPackage.schemaMain);
                        cf.Style.CreateNodes(new XmlHelperInstance(ws.NameSpaceManager, elem), "");
                        dxfsNode.AppendChild(elem);
                    }
                    else
                    {
                        ((ExcelConditionalFormattingRule)cf).DxfId = ix;
                    }
                }
            }
        }
        internal static void CopyDxfStylesTable(ExcelTable tblFrom, ExcelTable tblTo)
        {
            if (tblFrom.HeaderRowStyle.HasValue) tblTo.HeaderRowStyle = (ExcelDxfStyle)tblFrom.HeaderRowStyle.Clone();
            if (tblFrom.HeaderRowBorderStyle.HasValue) tblTo.HeaderRowBorderStyle = (ExcelDxfBorderBase)tblFrom.HeaderRowBorderStyle.Clone();
            if (tblFrom.DataStyle.HasValue) tblTo.DataStyle = (ExcelDxfStyle)tblFrom.DataStyle.Clone();
            if (tblFrom.TableBorderStyle.HasValue) tblTo.TableBorderStyle = (ExcelDxfBorderBase)tblFrom.TableBorderStyle.Clone();
            if (tblFrom.TotalsRowStyle.HasValue) tblTo.TotalsRowStyle = (ExcelDxfStyle)tblFrom.TotalsRowStyle.Clone();
            for (int c = 0; c < tblFrom.Columns.Count; c++)
            {
                var colFrom = tblFrom.Columns[c];
                var colTo = tblTo.Columns[c];
                if (colFrom.HeaderRowStyle.HasValue) colTo.HeaderRowStyle = (ExcelDxfStyle)colFrom.HeaderRowStyle.Clone();
                if (colFrom.DataStyle.HasValue) colTo.DataStyle = (ExcelDxfStyle)colFrom.DataStyle.Clone();
                if (colFrom.TotalsRowStyle.HasValue) colTo.TotalsRowStyle = (ExcelDxfStyle)colFrom.TotalsRowStyle.Clone();
            }
        }
    }
}
