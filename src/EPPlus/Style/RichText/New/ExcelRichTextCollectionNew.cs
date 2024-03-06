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

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Collection of Richtext objects
    /// </summary>
    public class ExcelRichTextCollectionNew : IEnumerable<ExcelRichTextNew>
    {
        List<ExcelRichTextNew> _list = new List<ExcelRichTextNew>();
        internal ExcelRangeBase _cells = null;
        internal ExcelWorkbook _wb;

        internal ExcelRichTextCollectionNew(XmlReader xr, ExcelWorkbook wb)
        {
            _wb = wb;
            while (xr.LocalName != "si" && xr.NodeType != XmlNodeType.EndElement) 
            {
                if (xr.LocalName == "r" && xr.NodeType == XmlNodeType.Element)
                {
                    XmlReaderHelper.ReadUntil(xr, "rPr", "t");
                    var item = new ExcelRichTextNew(this);
                    if (xr.LocalName == "rPr" && xr.NodeType == XmlNodeType.Element)
                    {
                        item.ReadrPr(xr);
                        xr.Read();
                    }
                    if (xr.LocalName == "t" && xr.NodeType == XmlNodeType.Element)
                    {
                        item.Text = xr.ReadElementContentAsString();
                    }
                    _list.Add(item);
                }
                xr.Read();
            }
        }

        internal ExcelRichTextCollectionNew(string s, ExcelWorkbook wb)
        {
            _wb = wb;
            var item = new ExcelRichTextNew(this);
            item.Text = s;
            _list.Add(item);
        }

        /// <summary>
        /// Collection containing the richtext objects
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        public ExcelRichTextNew this[int Index]
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
        /// <param name="NewParagraph">Adds a new paragraph before text. This will add a new line break.</param>
        /// <returns></returns>
        public ExcelRichTextNew Add(string Text, bool NewParagraph = false)
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
        public ExcelRichTextNew Insert(int index, string text)
        {
            return null;
        }

        internal void ConvertRichtext()
        {
            if (_cells == null) return;
            var isRt = _cells.Worksheet._flags.GetFlagValue(_cells._fromRow, _cells._fromCol, CellFlags.RichText);
            if (Count == 1 && isRt == false)
            {
                _cells.Worksheet._flags.SetFlagValue(_cells._fromRow, _cells._fromCol, true, CellFlags.RichText);
                var s = _cells.Worksheet.GetStyleInner(_cells._fromRow, _cells._fromCol);
                //var fnt = cell.Style.Font;
                var fnt = _cells.Worksheet.Workbook.Styles.GetStyleObject(s, _cells.Worksheet.PositionId, ExcelAddressBase.GetAddress(_cells._fromRow, _cells._fromCol)).Font;
                this[0].PreserveSpace = true;
                this[0].Bold = fnt.Bold;
                this[0].FontName = fnt.Name;
                this[0].Italic = fnt.Italic;
                this[0].Size = fnt.Size;
                this[0].UnderLine = fnt.UnderLine;

                int hex;
                if (fnt.Color.Rgb != "" && int.TryParse(fnt.Color.Rgb, NumberStyles.HexNumber, null, out hex))
                {
                    this[0].Color = Color.FromArgb(hex);
                }
            }
        }

        /// <summary>
        /// Clear the collection
        /// </summary>
        public void Clear()
        {
            _list.Clear();
            if (_cells != null)
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
            if (_cells != null && _list.Count == 0) _cells.SetIsRichTextFlag(false);
        }
        /// <summary>
        /// Removes an item
        /// </summary>
        /// <param name="Item"></param>
        public void Remove(ExcelRichTextNew Item)
        {
            _list.Remove(Item);
            if (_cells != null && _list.Count == 0) _cells.SetIsRichTextFlag(false);
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
                else
                {
                    this[0].Text = value;
                    for (int ix = 1; ix < Count; ix++)
                    {
                        RemoveAt(ix);
                    }
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
        #region IEnumerable<ExcelRichText> Members

        IEnumerator<ExcelRichTextNew> IEnumerable<ExcelRichTextNew>.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        internal string GetXML()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<si>");
            foreach(var item in _list)
            {
                item.GetXML(sb);
            }
            sb.Append("</si>");
            return sb.ToString();
        }

        #endregion
    }
}
