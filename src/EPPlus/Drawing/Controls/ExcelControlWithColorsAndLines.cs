/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    11/24/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/
using System.Xml;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Packaging;
namespace OfficeOpenXml.Drawing.Controls
{
    public abstract class ExcelControlWithColorsAndLines : ExcelControlWithText
    {
        internal ExcelControlWithColorsAndLines(ExcelDrawings drawings, XmlNode drawingNode, ControlInternal control, ZipPackagePart part, XmlDocument ctrlPropXml, ExcelGroupShape parent = null) :
            base(drawings, drawingNode, control, part, ctrlPropXml, parent)
        {

        }

        internal ExcelControlWithColorsAndLines(ExcelDrawings drawings, XmlElement drawNode) : base(drawings, drawNode)
        {
        }

        ExcelVmlDrawingFill _fill = null;
        public ExcelVmlDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelVmlDrawingFill(_drawings, _vml.NameSpaceManager, _vml.TopNode, _vml.SchemaNodeOrder);
                }
                return _fill;
            }
        }
        ExcelVmlDrawingBorder _border = null;
        public ExcelVmlDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelVmlDrawingBorder(_drawings, _vml.NameSpaceManager, _vml.TopNode, _vml.SchemaNodeOrder);
                }
                return _border;
            }
        }
        internal override void UpdateXml()
        {
            base.UpdateXml();
            Border.UpdateXml();
        }
    }
}
