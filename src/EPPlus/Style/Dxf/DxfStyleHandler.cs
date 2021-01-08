using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.Linq;
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
                    var item = new ExcelDxfStyleLimitedFont(styles.NameSpaceManager, x, styles);
                    styles.Dxfs.Add(item.Id, item);
                }
            }
        }
        internal static int CloneDxfStyle(ExcelWorkbook wb, int styleId)
        {
            var styles = wb.Styles;
            var copy = styles.Dxfs[styleId];
            var ix = styles.Dxfs.FindIndexById(copy.Id);
            if (ix < 0)
            {
                var parent = styles.GetNode(dxfsPath);
                var node = styles.TopNode.OwnerDocument.CreateElement("d:dxf", ExcelPackage.schemaMain);
                parent.AppendChild(node);
                node.InnerXml = copy._helper.TopNode.InnerXml;
                var dxf = new ExcelDxfStyleLimitedFont(styles.NameSpaceManager, node, styles);
                styles.Dxfs.Add(copy.Id, dxf);
                return styles.Dxfs.Count - 1;
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
        }

        private static void AddDxfNode(ExcelStyles styles, XmlNode dxfsNode, ExcelDxfStyle dxfStyle)
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
            }
        }

        private static void UpdateConditionalFormatting(ExcelWorksheet ws, ExcelStyleCollection<ExcelDxfStyle> dxfs, XmlNode dxfsNode)
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
