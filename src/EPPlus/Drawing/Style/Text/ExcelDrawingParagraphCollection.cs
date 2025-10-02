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
using OfficeOpenXml.Core.Worksheet.Fonts;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    public class ExcelDrawingParagraphCollection : XmlHelper, IEnumerable<ExcelDrawingParagraph>
    {
        IPictureRelationDocument _prd;
        string _path;
        Action _initXml = null;
        List<ExcelDrawingParagraph> _paragraphs = new List<ExcelDrawingParagraph>();
        internal Action<ExcelParagraphTextRunBase> _addCallback = null;
        internal Action<ExcelDrawingParagraph> _removeParagraphCallback = null;
        internal Action<ExcelParagraphTextRunBase> _removeTextRunCallback = null;
        internal ExcelDrawingParagraphCollection(IPictureRelationDocument prd, XmlNamespaceManager nsm, XmlNode topNode, string path, string[] schemaNodeOrder, Action initXml) : base(nsm, topNode)
        {
            _prd = prd;
            var rootNode = GetNode(path);
            _path = path;
            if (rootNode != null)
            {
                TopNode = rootNode;
                var pNodes = rootNode.SelectNodes("a:p", NameSpaceManager);
                foreach (XmlElement pn in pNodes)
                {
                    _paragraphs.Add(new ExcelDrawingParagraph(this, prd, NameSpaceManager, pn, schemaNodeOrder, initXml));
                }

                if (_paragraphs.Count == 0)
                {
                    CreateParagraphPlaceHolder();
                }
            }
            AddSchemaNodeOrder(schemaNodeOrder, ["rPr", "pPr", "t"]);
        }
        public int Count { get => _paragraphs.Count; }

        public ExcelDrawingParagraph this[int index]
        {
            get
            {
                return _paragraphs[index];
            }

        }
        /// <summary>
        /// Adds a new paragraph
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public ExcelDrawingParagraph Add(string text)
        {
            XmlNode pn = placeHolderNode;
            if (_paragraphs.Count == 0 && pn == null)
            {
                CreateTopNode();
            }

            if (pn == null)
            {
                pn = CreateNode("a:p", false, true);
            }

            var p = new ExcelDrawingParagraph(this, _prd, NameSpaceManager, pn, SchemaNodeOrder, _initXml);
            var tr = p.TextRuns.Add(text);
            _paragraphs.Add(p);
            //_addCallback?.Invoke(tr);
            if(placeHolderNode != null)
            {
                placeHolderNode = null;
            }
            return p;
        }

        XmlNode placeHolderNode = null;

        internal void CreateParagraphPlaceHolder()
        {
            if(placeHolderNode == null && _paragraphs.Count == 0)
            {
                var pn = CreateNode("a:p", false, true);
                placeHolderNode = pn;
            }
        }

        internal XmlNode CreateAndReturnParagraphPlaceHolder()
        {
            if(placeHolderNode == null)
            {
                CreateParagraphPlaceHolder();
            }
            return placeHolderNode;
        }
        /// <summary>
        /// Removes the item at the index from the collection
        /// </summary>
        /// <param name="index">The index</param>
        /// <exception cref="IndexOutOfRangeException"></exception>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= _paragraphs.Count)
            {
                throw new IndexOutOfRangeException("Paragraph index out of range.");
            }
            var pn = _paragraphs[index].TopNode;
            pn.ParentNode.RemoveChild(pn);
            _paragraphs.RemoveAt(index);
            _removeParagraphCallback?.Invoke(_paragraphs[index]);
        }
        /// <summary>
        /// Removes the item from the collection
        /// </summary>
        /// <param name="item">The item to remove</param>
        /// <exception cref="ArgumentException"></exception>
        public void Remove(ExcelDrawingParagraph item)
        {
            if (!_paragraphs.Contains(item))
            {
                throw new ArgumentException("Paragraph item does not exist in the collection");
            }
            var pn = item.TopNode;
            pn.ParentNode.RemoveChild(pn);
            _paragraphs.Remove(item);
            _removeParagraphCallback?.Invoke(item);
        }
        /// <summary>
        /// Gets the enumerator.
        /// </summary>
        /// <returns></returns>
        public IEnumerator<ExcelDrawingParagraph> GetEnumerator()
        {
            return _paragraphs.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public int IndexOf(ExcelDrawingParagraph item)
        {
            return _paragraphs.IndexOf(item);
        }
        /// <summary>
        /// Clear all paragraphs from the collection
        /// </summary>
        public void Clear()
        {
            for (int i = 0; i < _paragraphs.Count; i++)
            {
                var pn= _paragraphs[i].TopNode;
                pn.ParentNode.RemoveChild(pn);
                _removeParagraphCallback?.Invoke(_paragraphs[0]);
            }
            _paragraphs.Clear();
        }
        /// <summary>
        /// The text
        /// </summary>
        public string Text
        {
            get
            {
                if(Count==0) return "";
                StringBuilder sb = new StringBuilder();
                
                foreach(var p in _paragraphs)
                {
                    foreach(var tr in p.TextRuns)
                    {
                        sb.Append(tr.Text);
                    }
                    sb.Append(Environment.NewLine);
                }

                return sb.ToString().Substring(0, sb.Length - Environment.NewLine.Length);
            }
        }
        /// <summary>
        /// Returns true if the paragraph exists in the collection.
        /// </summary>
        /// <param name="item">The paragraph to check for</param>
        /// <returns>True if exists</returns>
        public bool Contains(ExcelDrawingParagraph item)
        {
            return _paragraphs.Contains(item);
        }
        /// <summary> 
        /// Creates the top nodes of the collection
        /// </summary>
        protected internal void CreateTopNode()
        {
            if (_paragraphs.Count == 0 && placeHolderNode == null)
            {
                if(GetNode(_path) == null)
                {
                    _initXml?.Invoke();
                    TopNode = CreateNode(_path);
                    CreateNode("a:bodyPr");
                    CreateNode("a:lstStyle");
                }
            }
        }

        internal void SetUpdateCallbacks(Action<ExcelParagraphTextRunBase> addCallback, Action<ExcelDrawingParagraph> removeParagraphCallback, Action<ExcelParagraphTextRunBase> removeTextRunCallback)
        {
            _addCallback = addCallback;
            _removeParagraphCallback = removeParagraphCallback;
            _removeTextRunCallback = removeTextRunCallback;            
        }
        internal RectBase GetSizeInPixels(double maxWidth, double maxHeight, string defaultText)
        {
            var rect = new RectBase();
            if (_paragraphs.Count == 0)
            {
                var ns = _prd.Package.Workbook.Styles.GetNormalStyle();
                var defFont = ns.Style.Font.GetMeasureFont();
                var t =_prd.Package.Settings.TextSettings.PrimaryTextMeasurer.MeasureText(defaultText, defFont); //TODO: use WrapMeasurer
            }

            foreach (var p in _paragraphs)
            {
                var pr = p.GetParagraphSizeInPixels(maxWidth, maxHeight);
                if(rect.Width < pr.Width) rect.Right = pr.Width;
                rect.Bottom += pr.Height;
                //foreach(var tr in p.TextRuns)
                //{
                //    tr.Text
                //    width += w;
                //    if (height < h) height = h;
                //}
            }
            return rect;
        }
        internal double GetHeightInPixels(ITextMeasurerWrap textMeasurer)
        {
            double totHeight = 0;
            foreach (var paragraph in _paragraphs)
            {
                //totHeight += paragraph.GetParagraphHeightInPixels(textMeasurer);
            }

            return totHeight;
            //textWidth = textHeight = 0;
            //var tm = _prd.Package.Settings.TextSettings.PrimaryTextMeasurer;
            //float lineWidth, lineHeight;
            //textWidth = textHeight = lineWidth = lineHeight = 0;
            //foreach (var p in _list)
            //{
            //    var fontName = string.IsNullOrEmpty(r.LatinFont) ? _defaultFont.LatinFont : r.LatinFont;
            //    if (fontName.StartsWith("+"))
            //    {
            //        var t = _drawing._drawings._package.Workbook.ThemeManager.GetOrCreateTheme();
            //        fontName = t.GetFontByCode(fontName);
            //    }
            //    var f = new MeasurementFont()
            //    {
            //        FontFamily = fontName,
            //        Size = r.Size <= 0 ? _defaultFont.Size : r.Size,
            //        Style = GetFontStyle(r)
            //    };
            //    var b = tm.MeasureText(r.Text, f);
            //    if (r.IsFirstInParagraph)
            //    {
            //        lineWidth = b.Width;
            //        lineHeight = b.Height;
            //    }
            //    else
            //    {
            //        lineWidth += b.Width;
            //        if (lineHeight < b.Height)
            //        {
            //            lineHeight = b.Height;
            //        }
            //    }

            //    if (r.IsLastInParagraph)
            //    {
            //        if (lineWidth > textWidth)
            //        {
            //            textWidth = lineWidth;
            //        }
            //        textHeight += lineHeight;
            //    }
            //}
        }

    }

}