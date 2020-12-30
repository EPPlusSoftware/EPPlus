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
                    ExcelDxfStyleConditionalFormatting item = new ExcelDxfStyleConditionalFormatting(styles.NameSpaceManager, x, styles);
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
                var dxf = new ExcelDxfStyleConditionalFormatting(styles.NameSpaceManager, node, styles);
                styles.Dxfs.Add(copy.Id, dxf);
                return styles.Dxfs.Count - 1;
            }
            else
            {
                return ix;
            }
        }

        const string dxfsPath = "d:styleSheet/d:dxfs";
        internal static void UpdateDxfXml(ExcelWorkbook wb)
        {
            //Set dxf styling for conditional Formatting
            var styles = wb.Styles;
            XmlNode dxfsNode = styles.TopNode.SelectSingleNode(dxfsPath, wb.NameSpaceManager);
            foreach (var ws in wb.Worksheets)
            {
                if (ws is ExcelChartsheet) continue;
                UpdateConditionalFormatting(ws, styles.Dxfs, dxfsNode);
                foreach(var pt in ws.PivotTables)
                {
                    if(pt.PivotAreaStyles!=null)
                    {
                        foreach(var pas in pt.PivotAreaStyles._list)
                        {
                            if(pas.Style.HasValue)
                            {
                                var ix = styles.Dxfs.FindIndexById(pas.Style.Id);
                                if (ix < 0)
                                {
                                    pas.Style.DxfId = styles.Dxfs.Count;
                                    styles.Dxfs.Add(pas.Style.Id, pas.Style);
                                    var elem = dxfsNode.OwnerDocument.CreateElement("dxf", ExcelPackage.schemaMain);
                                    pas.Style.CreateNodes(new XmlHelperInstance(ws.NameSpaceManager, elem), "");
                                    dxfsNode.AppendChild(elem);
                                }
                                else
                                {
                                    pas.Style.DxfId = ix;
                                }
                            }
                        }
                    }
                }
            }
            if (dxfsNode != null) (dxfsNode as XmlElement).SetAttribute("count", styles.Dxfs.Count.ToString());
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
