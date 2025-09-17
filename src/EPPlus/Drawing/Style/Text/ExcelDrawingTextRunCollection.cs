/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    9/11/2025         EPPlus Software AB       EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    public class ExcelDrawingTextRunCollection : XmlHelper, IEnumerable<ExcelParagraphTextRunBase>
    {
        List<ExcelParagraphTextRunBase> _textRuns;
        ExcelDrawingParagraph _paragraph;
        Action _initXml;
        internal ExcelDrawingTextRunCollection(ExcelDrawingParagraph paragraph, XmlNamespaceManager nsm, XmlNode topNode, Action initXml) : base(nsm, topNode)
        {
            _paragraph = paragraph;
            AddSchemaNodeOrder(_paragraph.SchemaNodeOrder, ["rPr", "pPr", "t"]);
            _initXml = initXml;
            _textRuns = new List<ExcelParagraphTextRunBase>();
            foreach (XmlElement node in topNode.SelectNodes("a:r|a:fld|a:br", nsm))
            {
                
                switch(node.LocalName)
                {
                    case "r":
                        _textRuns.Add(new ExcelParagraphTextRun(paragraph, nsm, node));
                        break;
                    case "fld":
                        _textRuns.Add(new ExcelParagraphTextField(paragraph, nsm, node));
                        break;
                    case "br":
                        _textRuns.Add(new ExcelParagraphLineBreak(paragraph, nsm, node));
                        break;

                }
            }
        }
        /// <summary>
        /// Number of items in the collection.
        /// </summary>
        public int Count { get => _textRuns.Count; }
        /// <summary>
        /// Return the text run at the index.
        /// </summary>
        /// <param name="index">The index</param>
        /// <returns></returns>
        public ExcelParagraphTextRunBase this[int index]
        {
            get
            {
                return _textRuns[index];
            }
        }

        /// <summary>
        /// Removes the item at the index from the collection
        /// </summary>
        /// <param name="index">The index</param>
        /// <exception cref="IndexOutOfRangeException"></exception>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= _textRuns.Count)
            {
                throw new IndexOutOfRangeException("Paragraph index out of range.");
            }
            var pn = _textRuns[index].TopNode;
            pn.ParentNode.RemoveChild(pn);
            _textRuns.RemoveAt(index);
        }
        /// <summary>
        /// Removes the item from the collection
        /// </summary>
        /// <param name="item">The item to remove</param>
        /// <exception cref="ArgumentException"></exception>
        public void Remove(ExcelParagraphTextRunBase item)
        {
            if (!_textRuns.Contains(item))
            {
                throw new ArgumentException("Paragraph item does not exist in the collection");
            }
            var pn = item.TopNode;
            pn.ParentNode.RemoveChild(pn);
            _textRuns.Remove(item);
        }
        /// <summary>
        /// Adds a rich text run with the text.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public ExcelParagraphTextRun Add(string text)
        {
            var rn=CreateNode("a:r", false, true);
            var txtRun = new ExcelParagraphTextRun(_paragraph, NameSpaceManager, rn);
            txtRun.Text = text;
            _textRuns.Add(txtRun);
            return txtRun;
        }
        internal ExcelParagraphTextRun Add(ExcelParagraphTextRun txtRun)
        {
            _textRuns.Add(txtRun);
            return txtRun;
        }
        /// <summary>
        /// Clear all text runs from the collection
        /// </summary>
        public void Clear()
        {
            for (int i = 0; i < _textRuns.Count; i++)
            {
                var pn = _textRuns[i].TopNode;
                pn.ParentNode.RemoveChild(pn);
            }
            _textRuns.Clear();
        }
        /// <summary>
        /// Returns true if the text run exists in the collection.
        /// </summary>
        /// <param name="item">The paragraph to check for</param>
        /// <returns>True if exists</returns>
        public bool Contains(ExcelParagraphTextRunBase item)
        {
            return _textRuns.Contains(item);
        }
        /// <summary>
        /// Returns the index in the collection of the supplied item.
        /// </summary>
        /// <param name="item">The item</param>
        /// <returns>The index in the collection.</returns>
        public int IndexOf(ExcelParagraphTextRunBase item)
        {
            return _textRuns.IndexOf(item);
        }
        IEnumerator<ExcelParagraphTextRunBase> IEnumerable<ExcelParagraphTextRunBase>.GetEnumerator()
        {
            return ((IEnumerable<ExcelParagraphTextRunBase>)_textRuns).GetEnumerator();
        }

        public IEnumerator<ExcelParagraphTextRunBase> GetEnumerator()
        {
            return _textRuns.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}