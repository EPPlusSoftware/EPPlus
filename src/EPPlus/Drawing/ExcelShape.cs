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
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// An Excel shape.
    /// </summary>
    public sealed class ExcelShape : ExcelShapeBase
    {
        internal ExcelShape(ExcelDrawings drawings, XmlNode node, ExcelGroupShape shape=null, DrawingsCollectionType DrawingsType = DrawingsCollectionType.excel) :
            base(drawings, node, NamespacePrefixes[(int)DrawingsType] + ":sp", NamespacePrefixes[(int)DrawingsType] + ":nvSpPr/"+ NamespacePrefixes[(int)DrawingsType] + ":cNvPr", shape, DrawingsType)
        {
        }
        internal ExcelShape(ExcelDrawings drawings, XmlNode node, eShapeStyle style, DrawingsCollectionType DrawingsType = DrawingsCollectionType.excel) :
            base(drawings, node, NamespacePrefixes[(int)DrawingsType]+":sp", NamespacePrefixes[(int)DrawingsType]+":nvSpPr/" + NamespacePrefixes[(int)DrawingsType]+":cNvPr")
        {
            XmlElement shapeNode = CreateShapeNode();
            shapeNode.InnerXml = ShapeStartXml();
            switch(DrawingsType)
            {
                case DrawingsCollectionType.chart:
                    node.AppendChild(shapeNode.OwnerDocument.CreateElement("cdr", "clientData", ExcelPackage.schemaChartDrawing));
                    break;
                case DrawingsCollectionType.excel:
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
            xml.AppendFormat("<{2}:nvSpPr><{2}:cNvPr id=\"{0}\" name=\"{1}\" /><{2}:cNvSpPr /></{2}:nvSpPr><{2}:spPr><a:prstGeom prst=\"rect\"><a:avLst /></a:prstGeom></{2}:spPr><{2}:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\" /></a:schemeClr></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\" /></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\" /></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\" /></a:fontRef></{2}:style><{2}:txBody><a:bodyPr vertOverflow=\"clip\" rtlCol=\"0\" anchor=\"ctr\" /><a:lstStyle /><a:p></a:p></{2}:txBody>", _id, Name, NamespacePrefixes[(int)_drawings.DrawingsType]);
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
    }
}
