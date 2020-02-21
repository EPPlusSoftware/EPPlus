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
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing;
using System.Drawing;
using System.Linq;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// A collection of Paragraph objects
    /// </summary>
    public class ExcelParagraphCollection : XmlHelper, IEnumerable<ExcelParagraph>
    {
        List<ExcelParagraph> _list = new List<ExcelParagraph>();
        private readonly ExcelDrawing _drawing;
        private readonly string _path;
        private readonly List<XmlElement> _paragraphs=new List<XmlElement>();
        internal ExcelParagraphCollection(ExcelDrawing drawing, XmlNamespaceManager ns, XmlNode topNode, string path, string[] schemaNodeOrder) :
            base(ns, topNode)
        {
            _drawing = drawing;
            AddSchemaNodeOrder(schemaNodeOrder, new string[] { "strRef","rich", "f", "strCache", "bodyPr", "lstStyle", "p", "ptCount","pt","pPr", "lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont","buNone", "buAutoNum", "buChar","buBlip", "tabLst","defRPr", "r","br","fld" ,"endParaRPr" });

            _path = path;
            var par = (XmlElement)TopNode.SelectSingleNode(path, NameSpaceManager);
            _paragraphs.Add(par);
            var nl = par.SelectNodes("a:r", NameSpaceManager);
            if (nl != null)
            {
                foreach (XmlNode n in nl)
                {
                    if (_list.Count==0 || n.ParentNode!=_list[_list.Count-1].TopNode.ParentNode)
                    {
                        _paragraphs.Add((XmlElement)n.ParentNode);
                    }
                    _list.Add(new ExcelParagraph(drawing._drawings, ns, n, "",schemaNodeOrder));
                }
            }
        }
        /// <summary>
        /// The indexer for this collection
        /// </summary>
        /// <param name="Index">The index</param>
        /// <returns></returns>
        public ExcelParagraph this[int Index]
        {
            get
            {
                return _list[Index];
            }
        }
        /// <summary>
        /// Number of items in the collection
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
        /// <param name="NewParagraph">This will be a new line. Is ignored for first item added to the collection</param>
        /// <returns></returns>
        public ExcelParagraph Add(string Text, bool NewParagraph=false)
        {
            XmlDocument doc;
            if (TopNode is XmlDocument)
            {
                doc = TopNode as XmlDocument;
            }
            else
            {
                doc = TopNode.OwnerDocument;
            }
            XmlNode parentNode;
            if(NewParagraph && _list.Count!=0)
            {
                parentNode = CreateNode(_path, false, true);
                _paragraphs.Add((XmlElement)parentNode);
            }
            else if(_paragraphs.Count > 1)
            {
                parentNode = _paragraphs[_paragraphs.Count - 1];
            }
            else 
            {
                parentNode = CreateNode(_path);
                _paragraphs.Add((XmlElement)parentNode);
            }

            var node = doc.CreateElement("a", "r", ExcelPackage.schemaDrawings);
            parentNode.AppendChild(node);
            var childNode = doc.CreateElement("a", "rPr", ExcelPackage.schemaDrawings);
            node.AppendChild(childNode);
            var rt = new ExcelParagraph(_drawing._drawings, NameSpaceManager, node, "", SchemaNodeOrder);
            var normalStyle = _drawing._drawings.Worksheet.Workbook.Styles.GetNormalStyle();
            if (normalStyle == null)
            {
                rt.ComplexFont = "Calibri";
                rt.LatinFont = "Calibri";
            }
            else
            {
                rt.LatinFont = normalStyle.Style.Font.Name;
                rt.ComplexFont = normalStyle.Style.Font.Name;
            }
            rt.Size = 11;

            rt.Text = Text;
            _list.Add(rt);
            return rt;
        }
        /// <summary>
        /// Removes all items in the collection
        /// </summary>
        public void Clear()
        {
            for (int ix = 0 ; ix < _paragraphs.Count; ix++)
            {
                _paragraphs[ix].ParentNode.RemoveChild(_paragraphs[ix]);
            }
            _list.Clear();
            _paragraphs.Clear();
        }
        /// <summary>
        /// Remove the item at the specified index
        /// </summary>
        /// <param name="Index">The index</param>
        public void RemoveAt(int Index)
        {
            var node = _list[Index].TopNode;
            while (node != null && node.Name != "a:r")
            {
                node = node.ParentNode;
            }
            node.ParentNode.RemoveChild(node);
            _list.RemoveAt(Index);
        }
        /// <summary>
        /// Remove the specified item
        /// </summary>
        /// <param name="Item">The item</param>
        public void Remove(ExcelRichText Item)
        {
            TopNode.RemoveChild(Item.TopNode);
        }
        /// <summary>
        /// The full text 
        /// </summary>
        public string Text
        {
            get
            {
                StringBuilder sb = new StringBuilder();
                foreach (var item in _list)
                {
                    if (item.IsLastInParagraph)
                    {
                        sb.AppendLine(item.Text);
                    }
                    else
                    {
                        sb.Append(item.Text);
                    }
                }
                if (sb.Length > 2) sb.Remove(sb.Length - 2, 2); //Remove last crlf
                return sb.ToString();
            }
            set
            {
                if (Count == 0)
                {
                    Add(value);
                }
                else
                {
                    this[0].Text = value;
                    for (int ix = _list.Count-1; ix > 0; ix--)
                    {
                        RemoveAt(ix);
                    }
                }
            }
        }
        #region IEnumerable<ExcelRichText> Members

        IEnumerator<ExcelParagraph> IEnumerable<ExcelParagraph>.GetEnumerator()
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
