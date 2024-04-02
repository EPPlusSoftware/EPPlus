/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class BorderXml : IBorder
    {
        BorderItemXml _top;
        BorderItemXml _bottom;
        BorderItemXml _left;
        BorderItemXml _right;

        public bool HasValue
        {
            get;
        }

        internal BorderXml(ExcelBorderXml border)
        {
            HasValue = !string.IsNullOrEmpty(border.Id);
            _top = new BorderItemXml(border.Top);
            _bottom = new BorderItemXml(border.Bottom);
            _left = new BorderItemXml(border.Left);
            _right = new BorderItemXml(border.Right);
        }

        public IBorderItem Top { get { return _top; } }

        public IBorderItem Bottom { get { return _bottom; } }

        public IBorderItem Left { get { return _left; } }

        public IBorderItem Right { get { return _right; } }
    }
}
