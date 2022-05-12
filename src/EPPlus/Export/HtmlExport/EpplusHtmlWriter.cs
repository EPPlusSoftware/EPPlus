/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/16/2020         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
    internal partial class EpplusHtmlWriter : HtmlWriterBase
    {
        internal EpplusHtmlWriter(Stream stream, Encoding encoding, Dictionary<string, int> styleCache) : base(stream, encoding, styleCache)
        {
        }

        private readonly Stack<string> _elementStack = new Stack<string>();
        private readonly List<EpplusHtmlAttribute> _attributes = new List<EpplusHtmlAttribute>();

        public void AddAttribute(string attributeName, string attributeValue)
        {
            Require.Argument(attributeName).IsNotNullOrEmpty("attributeName");
            Require.Argument(attributeValue).IsNotNullOrEmpty("attributeValue");
            _attributes.Add(new EpplusHtmlAttribute { AttributeName = attributeName, Value = attributeValue });
        }
        public void RenderBeginTag(string elementName, bool closeElement = false)
        {
            _newLine = false;
            // avoid writing indent characters for a hyperlinks or images inside a td element
            if(elementName != HtmlElements.A && elementName != HtmlElements.Img)
            {
                WriteIndent();
            }
            _writer.Write($"<{elementName}");
            foreach (var attribute in _attributes)
            {
                _writer.Write($" {attribute.AttributeName}=\"{attribute.Value}\"");
            }
            _attributes.Clear();

            if (closeElement)
            {
                _writer.Write("/>");
                _writer.Flush();
            }
            else
            {
                _writer.Write(">");
                _elementStack.Push(elementName);
            }
        }

        public void RenderEndTag()
        {
            if (_newLine)
            {
                WriteIndent();
            }

            var elementName = _elementStack.Pop();
            _writer.Write($"</{elementName}>");
            _writer.Flush();
        }

        internal void SetClassAttributeFromStyle(ExcelRangeBase cell, bool isHeader, HtmlExportSettings settings, string additionalClasses)
        {            
            string cls = string.IsNullOrEmpty(additionalClasses) ? "" : additionalClasses;
            int styleId = cell.StyleID;
            ExcelStyles styles = cell.Worksheet.Workbook.Styles;
            if (styleId < 0 || styleId >= styles.CellXfs.Count)
            {
                return;
            }

            var xfs = styles.CellXfs[styleId];
            var styleClassPrefix = settings.StyleClassPrefix;
            if (settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.CellDataType &&
               xfs.HorizontalAlignment == ExcelHorizontalAlignment.General)
            {
                if (ConvertUtil.IsNumericOrDate(cell.Value))
                {
                    cls = $"{styleClassPrefix}ar";
                }
                else if (isHeader)
                {
                    cls = $"{styleClassPrefix}al";
                }
            }

            if (styleId == 0 || HasStyle(xfs) == false)
            {
                if (string.IsNullOrEmpty(cls) == false)
                    AddAttribute("class", cls);
                return;
            }

            string key = GetStyleKey(xfs);

            var ma = cell.Worksheet.MergedCells[cell._fromRow, cell._fromCol];
            if (ma != null)
            {
                var address = new ExcelAddressBase(ma);
                var bottomStyleId = cell.Worksheet._values.GetValue(address._toRow, address._fromCol)._styleId;
                var rightStyleId = cell.Worksheet._values.GetValue(address._fromRow, address._toCol)._styleId;
                key += bottomStyleId + "|" + rightStyleId;
            }

            int id;
            if (_styleCache.ContainsKey(key))
            {
                id = _styleCache[key];
            }
            else
            {
                id = _styleCache.Count + 1;
                _styleCache.Add(key, id);
            }
            cls += $" {styleClassPrefix}{settings.CellStyleClassName}{id}";
            AddAttribute("class", cls.Trim());
        }
    }
}
