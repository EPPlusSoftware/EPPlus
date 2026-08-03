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
using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.Svg;
using EPPlus.Export.Utils;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Renderer;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Text;
using System.Xml;
using static Microsoft.IO.RecyclableMemoryStreamManager;
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// An Excel shape.
    /// </summary>
    public sealed class ExcelShape : ExcelShapeBase
    {
        internal ExcelShape(ExcelDrawings drawings, XmlNode node, ExcelGroupShape shape=null, DrawingsCollectionType collectionType = DrawingsCollectionType.Worksheet) :
            base(drawings, node, NamespacePrefixes[(int)collectionType] + ":sp", NamespacePrefixes[(int)collectionType] + ":nvSpPr/" + NamespacePrefixes[(int)collectionType] + ":cNvPr", shape, collectionType)
        {
            if (collectionType == DrawingsCollectionType.Chart)
            {
                var offNode = (XmlElement)node.SelectSingleNode("cdr:sp/cdr:spPr/a:xfrm/a:off", NameSpaceManager);
                var extNode = (XmlElement)node.SelectSingleNode("cdr:sp/cdr:spPr/a:xfrm/a:ext", NameSpaceManager);
                _frmXPosition = new ExcelDrawingCoordinate(drawings.NameSpaceManager, offNode);
                _frmXSize = new ExcelDrawingSize(drawings.NameSpaceManager, extNode);
            }
        }
        internal ExcelShape(ExcelDrawings drawings, XmlNode node, eShapeStyle style, DrawingsCollectionType collectionType = DrawingsCollectionType.Worksheet) :
            base(drawings, node, NamespacePrefixes[(int)collectionType] + ":sp", NamespacePrefixes[(int)collectionType] + ":nvSpPr/" + NamespacePrefixes[(int)collectionType] + ":cNvPr", null, collectionType)
        {
            if (collectionType == DrawingsCollectionType.Chart)
            {
                node.OwnerDocument.DocumentElement.SetAttribute("xmlns:cdr", ExcelPackage.schemaChartDrawing);
                node.OwnerDocument.DocumentElement.SetAttribute("xmlns:a", ExcelPackage.schemaDrawings);
            }
            XmlElement shapeNode = CreateShapeNode();
            shapeNode.InnerXml = ShapeStartXml();
            switch(collectionType)
            {
                case DrawingsCollectionType.Chart:
                    int x = (int)(_drawings._screenWidth * EMU_PER_PIXEL * (From.X));
                    int y = (int)(_drawings._screenHeight * EMU_PER_PIXEL * (From.Y));
                    int cx = (int)(_drawings._screenWidth * EMU_PER_PIXEL * (To.X - From.X));
                    int cy = (int)(_drawings._screenHeight * EMU_PER_PIXEL * (To.Y - From.Y));
                    XmlElement xFrmNode = GetXfrmNode(shapeNode);
                    if (xFrmNode.ChildNodes.Count == 0)
                    {
                        CreateNode(xFrmNode, "a:off");
                        CreateNode(xFrmNode, "a:ext");
                    }
                    var offNode = (XmlElement)xFrmNode.SelectSingleNode("a:off", NameSpaceManager);
                    offNode.SetAttribute("x", x.ToString());
                    offNode.SetAttribute("y", y.ToString());
                    var extNode = (XmlElement)xFrmNode.SelectSingleNode("a:ext", NameSpaceManager);
                    extNode.SetAttribute("cx", cx.ToString());
                    extNode.SetAttribute("cy", cy.ToString());
                    _frmXPosition = new ExcelDrawingCoordinate(drawings.NameSpaceManager, offNode);
                    _frmXSize = new ExcelDrawingSize(drawings.NameSpaceManager, extNode);
                    break;
                case DrawingsCollectionType.Worksheet:
                default:
                    node.AppendChild(shapeNode.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings));
                    break;
            }
            Style = style;
        }

        #region "Private Methods"
        private string ShapeStartXml()
        {
            StringBuilder xml = new StringBuilder();
            xml.AppendFormat("<{2}:nvSpPr><{2}:cNvPr id=\"{0}\" name=\"{1}\" /><{2}:cNvSpPr /></{2}:nvSpPr><{2}:spPr><a:xfrm/><a:prstGeom prst=\"rect\"><a:avLst /></a:prstGeom></{2}:spPr><{2}:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\" /></a:schemeClr></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\" /></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\" /></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\" /></a:fontRef></{2}:style><{2}:txBody><a:bodyPr vertOverflow=\"clip\" rtlCol=\"0\" anchor=\"ctr\" /><a:lstStyle /></{2}:txBody>", Id, Name, NamespacePrefixes[(int)_drawings._collectionType]);
            return xml.ToString();
        }

        #endregion
        internal override void DeleteMe()
        {
            if (Fill.Style == eFillStyle.BlipFill)
            {
                    IPictureContainer container = Fill.BlipFill;
                _drawings._package.PictureStore.RemoveImage(container.ImageHash, container);
            }
            base.DeleteMe();
        }
        public string ToSvg(SvgRenderOptions options)
        {
            var sr = new ShapeRenderer(this);           
            var sb = new StringBuilder();
            var svg = new SvgShapeRenderer(this.GetBoundingBox(), sb, options);
            svg.Render(sr.RenderItems);
            return sb.ToString();
        }

    }
}
