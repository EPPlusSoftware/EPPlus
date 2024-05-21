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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Drawing;
using System.Globalization;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Collection of Richtext objects
    /// </summary>
    public class ExcelRichTextCollection : IEnumerable<ExcelRichText>
    {
        List<ExcelRichText> _list = new List<ExcelRichText>();
        internal ExcelRangeBase _cells = null;
        internal ExcelWorkbook _wb;
        internal bool _isComment=false;
        internal ExcelRichTextCollection(ExcelWorkbook wb, ExcelRangeBase cells)
        {
            _wb = wb;
            _cells = cells;
			_cells._worksheet._flags.SetFlagValue(_cells._fromRow, _cells._fromCol, true, CellFlags.RichText);
		}

		internal ExcelRichTextCollection(string s, ExcelRangeBase cells)
        {
            _wb = cells._workbook;
            _cells = cells;
            if (!string.IsNullOrEmpty(s))
            {
                Add(s);
            }
        }

        internal ExcelRichTextCollection(ExcelRichTextCollection rtc, ExcelRangeBase cells)
        {
            _wb = cells._workbook;
            _cells = cells;
            foreach(var item in rtc._list)
            {
                _list.Add(new ExcelRichText(item, this));
            }
        }

        internal ExcelRichTextCollection(XmlReader xr, ExcelWorkbook wb)
        {
            _wb = wb;
            while (xr.LocalName != "si" && xr.NodeType != XmlNodeType.EndElement && xr.EOF==false) 
            {
                if (xr.LocalName == "r" && xr.NodeType == XmlNodeType.Element)
                {
                    XmlReaderHelper.ReadUntil(xr, "rPr", "t");
                    ExcelRichText item = new ExcelRichText(xr, this);
                    _list.Add(item);
                }
                xr.Read();
            }
        }

        internal ExcelRichTextCollection(XmlNamespaceManager ns, XmlNode textElem, ExcelRangeBase cells)
        {
            _wb = cells._workbook;
            _cells= cells;
            _isComment = true;

			foreach (XmlNode rElement in textElem.ChildNodes)
            {
                if(rElement.LocalName == "r")
                {
                    var t = rElement.SelectSingleNode("d:t", ns);
                    var rt = new ExcelRichText(ConvertUtil.ExcelDecodeString(t.InnerText), this);

                    rt.Bold = XmlHelper.GetRichTextPropertyBool(rElement.SelectSingleNode("d:rPr/d:b", ns));
                    rt.Italic = XmlHelper.GetRichTextPropertyBool(rElement.SelectSingleNode("d:rPr/d:i", ns));
                    rt.Strike = XmlHelper.GetRichTextPropertyBool(rElement.SelectSingleNode("d:rPr/d:strike", ns));
                    rt.UnderLineType = XmlHelper.GetRichTextPropertyUnderlineType(rElement.SelectSingleNode("d:rPr/d:u", ns), out bool underline);
                    rt.UnderLine = underline;
                    rt.VerticalAlign = XmlHelper.GetRichTextPropertyVerticalAlignmentFont(rElement.SelectSingleNode("d:rPr/d:vertAlign", ns));
                    rt.Size = XmlHelper.GetRichTextProperyFloat(rElement.SelectSingleNode("d:rPr/d:sz", ns));
                    rt.FontName = XmlHelper.GetRichTextPropertyString(rElement.SelectSingleNode("d:rPr/d:rFont", ns));
                    rt.ColorSettings = XmlHelper.GetRichTextPropertyColor(rElement.SelectSingleNode("d:rPr/d:color", ns), rt);
                    rt.Charset = XmlHelper.GetRichTextPropertyInt(rElement.SelectSingleNode("d:rPr/d:charset", ns));
                    rt.Family = XmlHelper.GetRichTextPropertyInt(rElement.SelectSingleNode("d:rPr/d:family", ns));
                    _list.Add(rt);
                }
            }
        }

        /// <summary>
        /// Collection containing the richtext objects
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        public ExcelRichText this[int Index]
        {
            get
            {
                var item = _list[Index];
                return item;
            }
        }
        /// <summary>
        /// Items in the list
        /// </summary>
        public int Count
        {
            get
            {
                return _list.Count;
            }
        }
        /// <summary>
        /// Add a rich text string
        /// </summary>
        /// <param name="Text">The text to add</param>
        /// <param name="NewParagraph">Adds a new paragraph after the <paramref name="Text"/>. This will add a new line break.</param>
        /// <returns></returns>
        public ExcelRichText Add(string Text, bool NewParagraph = false)
        {
            if (NewParagraph) Text += "\n";
            return Insert(_list.Count, Text);
        }

        /// <summary>
        /// Insert a rich text string at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index at which rich text should be inserted.</param>
        /// <param name="text">The text to insert.</param>
        /// <returns></returns>
        public ExcelRichText Insert(int index, string text)
        {
            if (text == null) throw new ArgumentException("Text can't be null", "text");
            var rt = new ExcelRichText(text, this);
            rt.PreserveSpace = true;
            int prevIndex = 0;
            if(index > _list.Count)
            {
                prevIndex = _list.Count - 1;
            }
            else
            {
                prevIndex = index - 1;
            }
            if(_list.Count > 0)
            {
                var prevRT = _list[prevIndex];
                rt.Bold = prevRT.Bold;
                rt.Italic = prevRT.Italic;
                rt.Strike = prevRT.Strike;
                rt.UnderLineType = prevRT.UnderLineType;
                rt.VerticalAlign = prevRT.VerticalAlign;
                rt.Size = prevRT.Size;
                rt.FontName = prevRT.FontName;
                rt.Charset = prevRT.Charset;
                rt.Family = prevRT.Family;
                rt.ColorSettings = prevRT.ColorSettings.Clone();
                rt.PreserveSpace = prevRT.PreserveSpace;
            }
            else if(_cells == null)
            {
                rt.FontName = "Calibri";
                rt.Size = 11;
            }
            else
            {
                var style = _cells.Offset(0, 0).Style;
                rt.FontName = style.Font.Name;
                rt.Size = style.Font.Size;
                rt.Bold = style.Font.Bold;
                rt.Italic = style.Font.Italic;
                rt.PreserveSpace = true;
                rt.UnderLine = style.Font.UnderLine;
                int hex;
                var s = _cells.Worksheet.GetStyleInner(_cells._fromRow, _cells._fromCol);
                var fnt = _cells.Worksheet.Workbook.Styles.GetStyleObject(s, _cells.Worksheet.PositionId, ExcelAddressBase.GetAddress(_cells._fromRow, _cells._fromCol)).Font;
                if (fnt.Color.Rgb != "" && int.TryParse(fnt.Color.Rgb, NumberStyles.HexNumber, null, out hex))
                {
                    rt.Color = Color.FromArgb(hex);
                }
                if (_isComment == false)
                {
                    _cells._worksheet._flags.SetFlagValue(_cells._fromRow, _cells._fromCol, true, CellFlags.RichText);
                    //_cells.SetIsRichTextFlag(true);
                }
            }
            _list.Insert(index, rt);
            return rt;
        }

        /// <summary>
        /// Clear the collection
        /// </summary>
        public void Clear()
        {
            _list.Clear();
            if (_cells != null && _isComment == false)
            {
				_cells.DeleteMe(_cells, false, true, true, true, false, true, false, false, false);
                _cells.SetIsRichTextFlag(false);
            }
        }
        /// <summary>
        /// Removes an item at the specific index
        /// </summary>
        /// <param name="Index"></param>
        public void RemoveAt(int Index)
        {
            _list.RemoveAt(Index);
            if (_cells != null && _list.Count == 0 && _isComment == false) _cells.SetIsRichTextFlag(false);
        }
        /// <summary>
        /// Removes an item
        /// </summary>
        /// <param name="Item"></param>
        public void Remove(ExcelRichText Item)
        {
            _list.Remove(Item);
            if (_cells != null && _list.Count == 0 && _isComment == false) _cells.SetIsRichTextFlag(false);
        }
        /// <summary>
        /// The text
        /// </summary>
        public string Text
        {
            get
            {
                StringBuilder sb = new StringBuilder();
                foreach (var item in _list)
                {
                    sb.Append(item.Text);
                }
                return sb.ToString();
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    Clear();
                }
                else if (Count == 0)
                {
                    Add(value);
                }
                else if (Count > 1)
                {
                    while(Count != 1)
                    {
                        RemoveAt(1);
                    }
                    this[0].Text = value;
                } 
                else if(Count == 1)
                {
                    this[0].Text = value;
                }
            }
        }
        /// <summary>
        /// Returns the rich text as a html string.
        /// </summary>
        public string HtmlText
        {
            get
            {
                var sb = new StringBuilder();
                foreach (var item in _list)
                {
                    item.WriteHtmlText(sb);
                }
                return sb.ToString();
            }
        }

        internal string GetXML()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var item in _list)
            {
                item.WriteRichTextAttributes(sb);
            }
            return sb.ToString();
        }

        #region IEnumerable<ExcelRichText> Members

        IEnumerator<ExcelRichText> IEnumerable<ExcelRichText>.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        #endregion
    }
}
