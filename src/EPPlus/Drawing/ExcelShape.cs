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
        internal ExcelShape(ExcelDrawings drawings, XmlNode node, ExcelGroupShape shape=null) :
            base(drawings, node, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", shape)
        {
        }
        internal ExcelShape(ExcelDrawings drawings, XmlNode node, eShapeStyle style) :
            base(drawings, node, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr")
        {
            XmlElement shapeNode = node.OwnerDocument.CreateElement("xdr", "sp", ExcelPackage.schemaSheetDrawings);
            shapeNode.SetAttribute("macro", "");
            shapeNode.SetAttribute("textlink", "");
            node.AppendChild(shapeNode);

            shapeNode.InnerXml = ShapeStartXml();
            node.AppendChild(shapeNode.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings));
            Style = style;
        }
        #region "Private Methods"
        private string ShapeStartXml()
        {
            StringBuilder xml = new StringBuilder();
            xml.AppendFormat("<xdr:nvSpPr><xdr:cNvPr id=\"{0}\" name=\"{1}\" /><xdr:cNvSpPr /></xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst /></a:prstGeom></xdr:spPr><xdr:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\" /></a:schemeClr></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\" /></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\" /></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\" /></a:fontRef></xdr:style><xdr:txBody><a:bodyPr vertOverflow=\"clip\" rtlCol=\"0\" anchor=\"ctr\" /><a:lstStyle /><a:p></a:p></xdr:txBody>", _id, Name);
            return xml.ToString();
        }
        #endregion
        internal override void DeleteMe()
        {
            if (Fill.Style == eFillStyle.BlipFill)
            {
                    IPictureContainer container = Fill.BlipFill;
                _drawings._package.PictureStore.RemoveImage(container.ImageHash, this);
            }
            base.DeleteMe();
        }
    }
}
