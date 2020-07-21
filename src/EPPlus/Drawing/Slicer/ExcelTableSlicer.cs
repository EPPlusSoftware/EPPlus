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
using OfficeOpenXml.Filter;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{
    /*
      <xsd:complexType name="CT_Slicer">
       <xsd:sequence>
         <xsd:element name="extLst" type="x:CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
       </xsd:sequence>
       <xsd:attribute name="name" type="x:ST_Xstring" use="required"/>
       <xsd:attribute ref="xr10:uid" use="optional"/>
       <xsd:attribute name="cache" type="x:ST_Xstring" use="required"/>
       <xsd:attribute name="caption" type="x:ST_Xstring" use="optional"/>
       <xsd:attribute name="startItem" type="xsd:unsignedInt" use="optional" default="0"/>
       <xsd:attribute name="columnCount" type="xsd:unsignedInt" use="optional" default="1"/>
       <xsd:attribute name="showCaption" type="xsd:boolean" use="optional" default="true"/>
       <xsd:attribute name="level" type="xsd:unsignedInt" use="optional" default="0"/>
       <xsd:attribute name="style" type="x:ST_Xstring" use="optional"/>
       <xsd:attribute name="lockedPosition" type="xsd:boolean" use="optional" default="false"/>
       <xsd:attribute name="rowHeight" type="xsd:unsignedInt" use="required"/>
     </xsd:complexType>
     */
    public class ExcelTableSlicer : ExcelSlicer<ExcelTableSlicerCache>
    {
        ExcelTable _table;
        ExcelSlicerXmlSource _xmlSource;
        internal ExcelTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent = null) : base(drawings, node, parent)
        {
            _ws = drawings.Worksheet;
            var slicerNode = _ws.SlicerXmlSources.GetSource(Name, eSlicerSourceType.Table, out _xmlSource);
            _slicerXmlHelper = XmlHelperFactory.Create(NameSpaceManager, slicerNode);
        }
        internal ExcelTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelTableColumn column) : base(drawings, node)
        {
            TableColumn = column;
            _table = column.Table;
            var name = drawings.Worksheet.Workbook.GetTableSlicerName(column.Name);
            CreateDrawing(name);
            SlicerName = name;

            Caption = column.Name;
            RowHeight = 19;
            CacheName = "Slicer_" + name.Replace(" ", "_");

            var cache = new ExcelTableSlicerCache(NameSpaceManager);
            cache.Init(column);
            _cache = cache;            
        }

        private void CreateDrawing(string name)
        {
            XmlElement graphFrame = TopNode.OwnerDocument.CreateElement("mc","AlternateContent", ExcelPackage.schemaMarkupCompatibility);
            graphFrame.SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);
            graphFrame.SetAttribute("xmlns:sle15", ExcelPackage.schemaSlicer);
            TopNode.AppendChild(graphFrame);
            graphFrame.InnerXml = string.Format("<mc:Choice Requires=\"sle15\"><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"{0}\" name=\"{2}\"><a:extLst><a:ext uri=\"{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}\"><a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{1}\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2010/slicer\"><sle:slicer xmlns:sle=\"http://schemas.microsoft.com/office/drawing/2010/slicer\" name=\"{2}\"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback xmlns=\"\"><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"{0}\" name=\"\"/><xdr:cNvSpPr><a:spLocks noTextEdit=\"1\"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"1200150\" y=\"2971800\"/><a:ext cx=\"1828800\" cy=\"2524125\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val=\"white\"/></a:solidFill><a:ln w=\"1\"><a:solidFill><a:prstClr val=\"green\"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\"/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\" sz=\"1100\"/><a:t>This shape represents a table slicer. Table slicers are not supported in this version of Excel.If the shape was modified in an earlier version of Excel, or if the workbook was saved in Excel 2007 or earlier, the slicer can't be used.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>", 
                _id,
                "{" + Guid.NewGuid().ToString() + "}",
                name);

            TopNode.AppendChild(TopNode.OwnerDocument.CreateElement("clientData", ExcelPackage.schemaSheetDrawings));

            _xmlSource = _ws.SlicerXmlSources.GetOrCreateSource(eSlicerSourceType.Table);            
            var node = _xmlSource.XmlDocument.CreateElement("slicer", ExcelPackage.schemaMainX14);
            _xmlSource.XmlDocument.DocumentElement.AppendChild(node);
            _slicerXmlHelper = XmlHelperFactory.Create(NameSpaceManager, node);

            var slNode = _ws.GetExtLstSubNode("{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}", "x14:slicerList");
            if (slNode == null)
            {
                _ws.CreateNode("d:extLst/d:ext", false, true);
                slNode = _ws.CreateNode("d:extLst/d:ext/x14:slicerList", false, true);
                ((XmlElement)slNode.ParentNode).SetAttribute("uri", "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}");

                var xh = XmlHelperFactory.Create(NameSpaceManager, slNode);
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
        /// This filter must be a value filter.
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
            if (_drawings.Worksheet.Workbook._tableSlicerNames.Contains(name))
            {
                return false;
            }
            _drawings.Worksheet.Workbook._tableSlicerNames.Add(name);
            return true;
        }
}
}

