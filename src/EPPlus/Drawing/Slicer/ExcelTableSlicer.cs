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
using OfficeOpenXml.Table;
using System;
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
        internal ExcelTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent = null) : base(drawings, node, parent)
        {
            _ws = drawings.Worksheet;
            var slicerNode = _ws.SlicerXmlSources.GetSource(Name, eSlicerSourceType.Table);
            _slicerXmlHelper = XmlHelperFactory.Create(NameSpaceManager, slicerNode);
        }
        internal ExcelTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelTableColumn column) : base(drawings, node)
        {
            _table = column.Table;
            CreateDrawing(column.Name);
        }

        private void CreateDrawing(string name)
        {
            XmlElement graphFrame = TopNode.OwnerDocument.CreateElement("mc","AlternateContent", ExcelPackage.schemaMarkupCompatibility);
            graphFrame.SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);
            graphFrame.SetAttribute("xmlns:sle15", ExcelPackage.schemaSlicer);
            TopNode.AppendChild(graphFrame);
            graphFrame.InnerXml = string.Format("<mc:Choice Requires=\"sle15\"><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"{0}\" name=\"\"><a:extLst><a:ext uri=\"{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}\"><a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{1}\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2010/slicer\"><sle:slicer xmlns:sle=\"http://schemas.microsoft.com/office/drawing/2010/slicer\" name=\"Name\"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback xmlns=\"\"><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"0\" name=\"\"/><xdr:cNvSpPr><a:spLocks noTextEdit=\"1\"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"1200150\" y=\"2971800\"/><a:ext cx=\"1828800\" cy=\"2524125\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val=\"white\"/></a:solidFill><a:ln w=\"1\"><a:solidFill><a:prstClr val=\"green\"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\"/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\" sz=\"1100\"/><a:t>This shape represents a table slicer. Table slicers are not supported in this version of Excel.If the shape was modified in an earlier version of Excel, or if the workbook was saved in Excel 2007 or earlier, the slicer can't be used.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>", 
                _id,
                "{" + Guid.NewGuid().ToString() + "}");

            TopNode.AppendChild(TopNode.OwnerDocument.CreateElement("clientData", ExcelPackage.schemaSheetDrawings));

            var source = _ws.SlicerXmlSources.GetOrCreateSource(eSlicerSourceType.Table);
            //LoadXmlSafe(slicerXml, GetStartXml(), Encoding.UTF8);

            //// save it to the package
            //Part = package.CreatePart(UriChart, ExcelPackage.contentTypeChartEx, _drawings._package.Compression);

            //StreamWriter streamChart = new StreamWriter(Part.GetStream(FileMode.Create, FileAccess.Write));
            //ChartXml.Save(streamChart);
            //streamChart.Close();
            //package.Flush();

            //var chartRelation = drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(drawings.UriDrawing, UriChart), Packaging.TargetMode.Internal, ExcelPackage.schemaChartExRelationships);
            //graphFrame.SelectSingleNode("mc:Choice/xdr:graphicFrame/a:graphic/a:graphicData/cx:chart", NameSpaceManager).Attributes["r:id"].Value = chartRelation.Id;
            //package.Flush();
            //_chartNode = ChartXml.SelectSingleNode("cx:chartSpace/cx:chart", NameSpaceManager);
            //_chartXmlHelper = XmlHelperFactory.Create(NameSpaceManager, _chartNode);
            //GetPositionSize();
        }

        private Stream GetStartXml()
        {
            throw new NotImplementedException();
        }
    }
}
