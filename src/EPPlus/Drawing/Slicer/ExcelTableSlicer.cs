/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/26/2020         EPPlus Software AB       EPPlus 5.3
 ******0*******************************************************************************************/
using OfficeOpenXml.Constants;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{

    public class ExcelTableSlicer : ExcelSlicer<ExcelTableSlicerCache>
    {
        ExcelSlicerXmlSource _xmlSource;
        internal ExcelTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent = null) : base(drawings, node, parent)
        {
            _ws = drawings.Worksheet;
            var slicerNode = _ws.SlicerXmlSources.GetSource(Name, eSlicerSourceType.Table, out _xmlSource);
            _slicerXmlHelper = XmlHelperFactory.Create(NameSpaceManager, slicerNode);
            
            _ws.Workbook.SlicerCaches.TryGetValue(CacheName, out ExcelSlicerCache cache);
            _cache = (ExcelTableSlicerCache)cache;

            TableColumn = GetTableColumn();
        }

        internal ExcelTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelTableColumn column) : base(drawings, node)
        {
            TableColumn = column;
            column.Slicer = this;
            var name = drawings.Worksheet.Workbook.GetSlicerName(column.Name);
            CreateDrawing(name);
            SlicerName = name;

            Caption = column.Name;
            RowHeight = 19;
            CacheName = "Slicer_" + name.Replace(" ", "_");

            var cache = new ExcelTableSlicerCache(NameSpaceManager);
            cache.Init(column, CacheName);
            _cache = cache;            
        }
        private ExcelTableColumn GetTableColumn()
        {
            foreach (var ws in _drawings.Worksheet.Workbook.Worksheets)
            {
                foreach (var t in ws.Tables)
                {
                    if (t.Id == Cache.TableId)
                    {
                        return t.Columns.Where(x => x.Id == Cache.ColumnId).Single();
                    }
                }
            }
            return null;
        }

        internal override void DeleteMe()
        {
            try
            {
                _xmlSource.Part.Package.DeletePart(_xmlSource.Uri);
                _xmlSource.Part.Package.DeletePart(Cache.Uri);
                TableColumn.Slicer = null;
            }
            catch (Exception ex)
            {
                throw (new InvalidDataException("EPPlus internal error when deleting the slicer.", ex));
            }

            base.DeleteMe();
        }

        private void CreateDrawing(string name)
        {
            XmlElement graphFrame = TopNode.OwnerDocument.CreateElement("mc","AlternateContent", ExcelPackage.schemaMarkupCompatibility);
            graphFrame.SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);
            graphFrame.SetAttribute("xmlns:sle15", ExcelPackage.schemaSlicer);
            TopNode.AppendChild(graphFrame);
            graphFrame.InnerXml = string.Format("<mc:Choice Requires=\"sle15\"><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"{0}\" name=\"{2}\"><a:extLst><a:ext uri=\"{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}\"><a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{1}\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2010/slicer\"><sle:slicer xmlns:sle=\"http://schemas.microsoft.com/office/drawing/2010/slicer\" name=\"{2}\"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback xmlns=\"\"><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"{0}\" name=\"{2}\"/><xdr:cNvSpPr><a:spLocks noTextEdit=\"1\"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"1200150\" y=\"2971800\"/><a:ext cx=\"1828800\" cy=\"2524125\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val=\"white\"/></a:solidFill><a:ln w=\"1\"><a:solidFill><a:prstClr val=\"green\"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\"/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\" sz=\"1100\"/><a:t>This shape represents a table slicer. Table slicers are not supported in this version of Excel.If the shape was modified in an earlier version of Excel, or if the workbook was saved in Excel 2007 or earlier, the slicer can't be used.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>", 
                _id,
                "{" + Guid.NewGuid().ToString() + "}",
                name);

            TopNode.AppendChild(TopNode.OwnerDocument.CreateElement("clientData", ExcelPackage.schemaSheetDrawings));

            _xmlSource = _ws.SlicerXmlSources.GetOrCreateSource(eSlicerSourceType.Table);            
            var node = _xmlSource.XmlDocument.CreateElement("slicer", ExcelPackage.schemaMainX14);
            _xmlSource.XmlDocument.DocumentElement.AppendChild(node);
            _slicerXmlHelper = XmlHelperFactory.Create(NameSpaceManager, node);

            var extNode = _ws.GetOrCreateExtLstSubNode(ExtLstUris.WorksheetSlicerTableUri, "x14");
            if (extNode.InnerXml=="")
            {
                extNode.InnerXml = "<x14:slicerList/>";
                var xh = XmlHelperFactory.Create(NameSpaceManager, extNode.FirstChild);
                var element = (XmlElement)xh.CreateNode("x14:slicer", false, true);
                element.SetAttribute("id", ExcelPackage.schemaRelationships, _xmlSource.Rel.Id);
            }

            GetPositionSize();
        }
        public ExcelTableColumn TableColumn
        {
            get;
        }
        /// <summary>
        /// The value filters for the slicer. This is the same filter as the filter for the table.
        /// This filter is a value filter.
        /// </summary>
        public ExcelValueFilterCollection FilterValues
        {
            get
            {
                var f=TableColumn.Table.AutoFilter.Columns[TableColumn.Position] as ExcelValueFilterColumn;
                if(f!=null)
                {
                    return f.Filters;
                }
                else
                {
                    return null;
                }
            }
        } 

        internal override bool CheckSlicerNameIsUnique(string name)
        {
            if (_drawings.Worksheet.Workbook._slicerNames.Contains(name))
            {
                return false;
            }
            _drawings.Worksheet.Workbook._slicerNames.Add(name);
            return true;
        }
}
}

