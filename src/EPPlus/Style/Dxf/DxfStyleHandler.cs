using OfficeOpenXml.ConditionalFormatting;
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
        internal static void Load(ExcelWorkbook wb, ExcelStyles styles)
        {
            //dxfsPath
            XmlNode dxfsNode = styles.GetNode(dxfsPath);
            if (dxfsNode != null)
            {
                foreach (XmlNode x in dxfsNode)
                {
                    var item = new ExcelDxfStyle(styles.NameSpaceManager, x, styles);
                    styles.Dxfs.Add(item.Id, item);
                }
            }
        }
        internal static int CloneDxfStyle(ExcelStyles stylesFrom, ExcelStyles stylesTo, int styleId)
        {
            var copy = stylesFrom.Dxfs[styleId];
            var ix = stylesTo.Dxfs.FindIndexById(copy.Id);
            if (ix < 0)
            {
                var parent = stylesTo.GetNode(dxfsPath);
                var node = stylesTo.TopNode.OwnerDocument.CreateElement("d:dxf", ExcelPackage.schemaMain);
                parent.AppendChild(node);
                node.InnerXml = copy._helper.TopNode.InnerXml;
                ExcelDxfStyle dxf;
                //if (copy is ExcelDxfStyleLimitedFont)
                //{
                //    dxf = new ExcelDxfStyleLimitedFont(stylesTo.NameSpaceManager, node, stylesTo);
                //}
                //else 
                //{
                    dxf = new ExcelDxfStyle(stylesTo.NameSpaceManager, node, stylesTo);
                //}
                stylesTo.Dxfs.Add(copy.Id, dxf);
                return stylesTo.Dxfs.Count - 1;
            }
            else
            {
                return ix;
            }
        }

        const string dxfsPath = "d:dxfs";
        internal static void UpdateDxfXml(ExcelWorkbook wb)
        {
            //Set dxf styling for conditional Formatting
            XmlNode dxfsNode = wb.Styles.TopNode.SelectSingleNode(dxfsPath, wb.NameSpaceManager);
            UpdateTableStyles(wb, wb.Styles, dxfsNode);
            UpdateDxfXmlWorksheet(wb, wb.Styles, dxfsNode);
            if (dxfsNode != null) (dxfsNode as XmlElement).SetAttribute("count", wb.Styles.Dxfs.Count.ToString());
        }

        private static void UpdateTableStyles(ExcelWorkbook wb, ExcelStyles styles, XmlNode dxfsNode)
        {
            foreach (var ts in styles.TableStyles)
            {
                foreach(var element in ts._dic.Values)
                {
                    AddDxfNode(styles, dxfsNode, element.Style);
                    if(element.Style.DxfId>=0)
                    {
                        element.CreateNode();
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
                tbl.HeaderRowDxfId = AddDxfNode(styles, dxfsNode, tbl.HeaderRowStyle);
                tbl.DataDxfId = AddDxfNode(styles, dxfsNode, tbl.DataStyle);
                tbl.TotalsRowDxfId = AddDxfNode(styles, dxfsNode, tbl.TotalsRowStyle);
                
                tbl.HeaderRowBorderDxfId = AddDxfBorderNode(styles, dxfsNode, tbl.HeaderRowBorderStyle);
                tbl.TableBorderDxfId = AddDxfBorderNode(styles, dxfsNode, tbl.TableBorderStyle);

                foreach (var column in tbl.Columns)
                {
                    column.HeaderRowDxfId = AddDxfNode(styles, dxfsNode, column.HeaderRowStyle);
                    column.DataDxfId = AddDxfNode(styles, dxfsNode, column.DataStyle);
                    column.TotalsRowDxfId = AddDxfNode(styles, dxfsNode, column.TotalsRowStyle);
                }
            }
        }

        private static void UpdateDxfXmlPivotTables(ExcelStyles styles, XmlNode dxfsNode, ExcelWorksheet ws)
        {
            foreach (var pt in ws.PivotTables)
            {
                if (pt.Styling != null)
                {
                    foreach (var pas in pt.Styling.Areas._list)
                    {
                        AddDxfNode(styles, dxfsNode, pas.Style);
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
                    styles.Dxfs.Add(borderStyle.Id, new ExcelDxfStyle(styles.NameSpaceManager, elem, styles) { Border = borderStyle });
                    return styles.Dxfs.Count - 1;
                }
                else
                {
                    return ix;
                }
            }
            return null;
        }
        private static int? AddDxfNode(ExcelStyles styles, XmlNode dxfsNode, ExcelDxfStyleBase dxfStyle)
        {
            if (dxfStyle.HasValue)
            {
                var ix = styles.Dxfs.FindIndexById(dxfStyle.Id);
                if (ix < 0)
                {
                    dxfStyle.DxfId = styles.Dxfs.Count;
                    styles.Dxfs.Add(dxfStyle.Id, dxfStyle);
                    var elem = dxfsNode.OwnerDocument.CreateElement("dxf", ExcelPackage.schemaMain);
                    dxfStyle.CreateNodes(new XmlHelperInstance(styles.NameSpaceManager, elem), "");
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
    }
}
