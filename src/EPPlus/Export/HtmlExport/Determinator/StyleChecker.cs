using OfficeOpenXml.Export.HtmlExport.Parsers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Determinator
{
    internal class StyleChecker
    {
        ExcelStyles _styles;
        ExcelXfs _style;
        StyleCache _cache;
        int _id = -1;

        internal int Id => _id;

        List<ExcelXfs> _styleList; 

        internal StyleChecker(ExcelXfs style, StyleCache cache, ExcelStyles styles)
        {
            _style = style;
            _cache = cache;
            _styles = styles;
            _styleList = new List<ExcelXfs>
            { _style };
        }

        internal string GetStyleKey()
        {
            var fbfKey = ((ulong)(uint)_style.FontId << 32 | (uint)_style.BorderId << 16 | (uint)_style.FillId);
            return fbfKey.ToString() + "|" + ((int)_style.HorizontalAlignment).ToString() + "|" + ((int)_style.VerticalAlignment).ToString() + "|" + _style.Indent.ToString() + "|" + _style.TextRotation.ToString() + "|" + (_style.WrapText ? "1" : "0");
        }

        internal bool IsAdded(int bottomStyleId = -1, int rightStyleId = -1)
        {
            var key = AttributeTranslator.GetStyleKey(_style);
            if (bottomStyleId > -1) key += bottomStyleId + "|" + rightStyleId;

            bool ret = _cache.IsAdded(key, out _id);

            return ret;
        }

        internal bool BorderStyleCheck(int borderIdBottom, int borderIdRight)
        {
            return (HasStyle() || borderIdBottom > 0 || borderIdRight > 0);
        }

        internal bool ShouldAdd()
        {
            return !IsAdded();
        }

        internal bool ShouldAddWithBorders(int bottomStyleId, int rightStyleId)
        {
            bool added = IsAdded(bottomStyleId, rightStyleId);

            if (added)
            {
                return false;
            }

            _styleList.Add(_styles.CellXfs[bottomStyleId]);
            _styleList.Add(_styles.CellXfs[rightStyleId]);

            return BorderStyleCheck(_styleList[1].BorderId, _styleList[2].BorderId);
        }

        internal List<ExcelXfs> GetStyleList()
        {
            return _styleList;
        }

        internal bool HasStyle()
        {
            return _style.FontId > 0 ||
                   _style.FillId > 0 ||
                   _style.BorderId > 0 ||
                   _style.HorizontalAlignment != ExcelHorizontalAlignment.General ||
                   _style.VerticalAlignment != ExcelVerticalAlignment.Bottom ||
                   _style.TextRotation != 0 ||
                   _style.Indent > 0 ||
                   _style.WrapText;
        }
    }
}
