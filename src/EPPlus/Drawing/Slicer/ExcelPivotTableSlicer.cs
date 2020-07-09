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
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{
    public class ExcelPivotTableSlicer : ExcelSlicer<ExcelPivotTableSlicerCache>
    {
        ExcelSlicerXmlSource _source;
        internal ExcelPivotTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent = null) : base(drawings, node, parent)
        {
            _ws = drawings.Worksheet;
            var slicerNode = _ws.SlicerXmlSources.GetSource(Name, eSlicerSourceType.PivotTable, out _source);
            _slicerXmlHelper = XmlHelperFactory.Create(NameSpaceManager, slicerNode);
        }
        public ExcelPivotTableCollection PivotTables
        {
            get;
        } = new ExcelPivotTableCollection();
        private void CreateDrawing()
        {
            XmlElement graphFrame = TopNode.OwnerDocument.CreateElement("mc", "AlternateContent", ExcelPackage.schemaMarkupCompatibility);
            graphFrame.SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);
            graphFrame.SetAttribute("xmlns:a14", ExcelPackage.schemaDrawings2010);
            TopNode.AppendChild(graphFrame);
            graphFrame.InnerXml = string.Format("<mc:Choice Requires=\"a14\"><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"{0}\" name=\"\"><a:extLst><a:ext uri=\"{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}\"><a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{{{1}}\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2010/slicer\"><sle:slicer xmlns:sle=\"http://schemas.microsoft.com/office/drawing/2010/slicer\" name=\"\"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback xmlns=\"\"><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"0\" name=\"{1}\"/><xdr:cNvSpPr><a:spLocks noTextEdit=\"1\"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"12506325\" y=\"3238500\"/><a:ext cx=\"1828800\" cy=\"2524125\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val=\"white\"/></a:solidFill><a:ln w=\"1\"><a:solidFill><a:prstClr val=\"green\"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\"/><a:lstStyle/><a:p><a:r><a:rPr lang=\"sv-SE\" sz=\"1100\"/><a:t>This shape represents a slicer. Slicers are supported in Excel 2010 or later. If the shape was modified in an earlier version of Excel, or if the workbook was saved in Excel 2003 or earlier, the slicer cannot be used.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>",
                _id,
                Guid.NewGuid().ToString());
            TopNode.AppendChild(TopNode.OwnerDocument.CreateElement("clientData", ExcelPackage.schemaSheetDrawings));

            var slicerXml = new XmlDocument
            {
                PreserveWhitespace = ExcelPackage.preserveWhitespace
            };
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
            GetPositionSize();
        }
    }
}
