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
using System.Globalization;
using OfficeOpenXml.Interfaces.Drawing.Text;
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
        private readonly float _defaultFontSize;
        private readonly ExcelTextFont _defaultFont;
        private readonly ExcelTextBody _textBody;
        internal ExcelParagraphCollection(ExcelTextBody tb,  ExcelDrawing drawing, XmlNamespaceManager ns, XmlNode topNode, string path, string[] schemaNodeOrder, float defaultFontSize =11) :
            base(ns, topNode)
        {
            _drawing = drawing;
            _textBody = tb;
            _defaultFontSize = defaultFontSize;
            AddSchemaNodeOrder(schemaNodeOrder, new string[] { "strRef","rich", "f", "strCache", "bodyPr", "lstStyle", "p", "ptCount","pt","pPr", "lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont","buNone", "buAutoNum", "buChar","buBlip", "tabLst","defRPr", "r","br","fld" ,"endParaRPr" });
            _defaultFont = new ExcelTextFontXml(drawing._drawings, ns, TopNode, path+ "/a:pPr/a:defRPr", schemaNodeOrder);
            _path = path;
            foreach(var p in tb.Paragraphs)
            {
                foreach(var tr in p.TextRuns)
                {
                    _list.Add(new ExcelParagraph(tr));
                }
            }
            tb.Paragraphs.SetUpdateCallbacks(AddParagraph, RemoveParagraph, RemoveTextRun);
            var paths = path.Split('/');
        }

        private void AddParagraph(ExcelParagraphTextRunBase tr)
        {
            bool inParagraph = false;
            for (int i=0;i < _list.Count;i++)
            {                
                var item = _list[i];
                if(item._textRun.Paragraph==tr.Paragraph)
                {
                    inParagraph = true;
                }
                else if(inParagraph)
                {
                    _list.Insert(i - 1, new ExcelParagraph(tr));
                    return;
                }
            }
            _list.Add(new ExcelParagraph(tr));
        }
        private void RemoveTextRun(ExcelParagraphTextRunBase textRun)
        {
            for(int i=0;i<Count;i++)
            {
                if (_list[i].IsTextRun(textRun))
                {
                    _list.RemoveAt(i);
                    break;
                }
            }
        }

        private void RemoveParagraph(ExcelDrawingParagraph paragraph)
        {
            for (int i = 0; i < Count; i++)
            {
                if (_list[i].IsInParagraph(paragraph))
                {
                    _list.RemoveAt(i--);
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
        /// <param name="NewParagraph">This will be a new line. </param>
        /// <returns></returns>
        public ExcelParagraph Add(string Text, bool NewParagraph=false)
        {
            ExcelDrawingParagraph p;
            if(NewParagraph || _textBody.Paragraphs.Count==0)
            {
                p = _textBody.Paragraphs.Add(Text);
            }
            else
            {
                p = _textBody.Paragraphs[_textBody.Paragraphs.Count - 1];
                p.TextRuns.Add(Text);
            }
            var item = new ExcelParagraph(p.TextRuns[p.TextRuns.Count - 1]);
            _list.Add(item);
            return item;

            //XmlDocument doc;
            //if (TopNode is XmlDocument)
            //{
            //    doc = TopNode as XmlDocument;
            //}
            //else
            //{
            //    doc = TopNode.OwnerDocument;
            //}
            //XmlNode parentNode;
            //if(NewParagraph && _list.Count!=0)
            //{
            //    parentNode = CreateNode(_path, false, true);
            //    _paragraphs.Add((XmlElement)parentNode);
            //    var p = _list[0].TopNode.ParentNode.ParentNode.SelectSingleNode("a:pPr", NameSpaceManager);
            //    if(p!=null)
            //    {
            //        parentNode.InnerXml = p.OuterXml;
            //    }                
            //}
            //else if(_paragraphs.Count > 1)
            //{
            //    parentNode = _paragraphs[_paragraphs.Count - 1];
            //}
            //else 
            //{                
            //    parentNode = CreateNode(_path);
            //    _paragraphs.Add((XmlElement)parentNode);
            //    var defNode = CreateNode(_path + "/a:pPr/a:defRPr");
            //    if (defNode.InnerXml == "")
            //    {
            //        ((XmlElement)defNode).SetAttribute("sz", (_defaultFontSize*100).ToString(CultureInfo.InvariantCulture));
            //        var normalStyle = _drawing._drawings.Worksheet.Workbook.Styles.GetNormalStyle();
            //        if (normalStyle == null)
            //            defNode.InnerXml = "<a:latin typeface=\"Calibri\" /><a:cs typeface=\"Calibri\" />";
            //        else
            //            defNode.InnerXml = $"<a:latin typeface=\"{normalStyle.Style.Font.Name}\"/><a:cs typeface=\"{normalStyle.Style.Font.Name}\"/>";
            //    }
            //}

            //var node = doc.CreateElement("a", "r", ExcelPackage.schemaDrawings);
            //parentNode.AppendChild(node);
            //var childNode = doc.CreateElement("a", "rPr", ExcelPackage.schemaDrawings);
            //node.AppendChild(childNode);
            //var rt = new ExcelParagraph(_drawing._drawings, NameSpaceManager, node, "", SchemaNodeOrder);
            //rt.Text = Text;
            //_list.Add(rt);
            //return rt;
        }
        /// <summary>
        /// Removes all items in the collection
        /// </summary>
        public void Clear()
        {
            //for (int ix = 0 ; ix < _paragraphs.Count; ix++)
            //{
            //    _paragraphs[ix].ParentNode?.RemoveChild(_paragraphs[ix]);
            //}
            _textBody.Paragraphs.Clear();
            _list.Clear();
            //_paragraphs.Clear();
        }
        /// <summary>
        /// Remove the item at the specified index
        /// </summary>
        /// <param name="Index">The index</param>
        public void RemoveAt(int Index)
        {
            Remove(_list[Index]);
            //var node = _list[Index].TopNode;
            //while (node != null && node.Name != "a:r")
            //{
            //    node = node.ParentNode;
            //}
            //node.ParentNode.RemoveChild(node);
            //_list.RemoveAt(Index);
        }
        /// <summary>
        /// Remove the specified item
        /// </summary>
        /// <param name="Item">The item</param>
        public void Remove(ExcelParagraph Item)
        {
            var p = Item._textRun.Paragraph;
            p.TextRuns.Remove(Item._textRun);
            if (p.TextRuns.Count == 0 && _textBody.Paragraphs.Count != 0)
            {
                _textBody.Paragraphs.Remove(p);
            }
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
                
                var ret = sb.ToString();
                if (ret.EndsWith(Environment.NewLine))
                {
                    //Remove last NewLine
                    return ret.Substring(0, ret.Length - Environment.NewLine.Length);
                }

                return ret;
            }
            set
            {
                if (_textBody.Paragraphs.Count == 0)
                {
                    Add(value);
                }
                else
                {
                    if (Count == 0)
                    {
                        _textBody.Paragraphs[0].TextRuns.Add(value);
                    }
                    else
                    {
                        this[0].Text = value;
                        for (int ix = _list.Count - 1; ix > 0; ix--)
                        {
                            RemoveAt(ix);
                        }
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

        //internal void UpdateXmlEndParagraphRunProperties()
        //{
        //    if (_list.Count > 1)
        //    {
        //        for (int i = 1; i < _list.Count; i++)
        //        {
        //            //Get the run node of previous paragraph if it exists
        //            var prevRunNode = _list[i-1].TopNode;
        //            if (prevRunNode != null)
        //            {
        //                var endParaNode = _list[i].TopNode.SelectSingleNode("../../a:endParaRPr", NameSpaceManager);

        //                if(endParaNode != null)
        //                {
        //                    endParaNode.ParentNode.RemoveChild(endParaNode);
        //                }

        //                endParaNode = _list[i].CreateNode("../../a:endParaRPr");

        //                foreach(XmlAttribute attribute in prevRunNode.Attributes)
        //                {
        //                    endParaNode.Attributes.Append((XmlAttribute)attribute.Clone());
        //                }

        //                foreach(XmlNode childnode in prevRunNode.ChildNodes)
        //                {
        //                    endParaNode.AppendChild(childnode.Clone());
        //                }
        //            }
        //        }
        //    }
        //}

        internal void GetHeightInPixels(out float textWidth, out float textHeight)
        {
            var tm = _drawing._drawings._package.Settings.TextSettings.PrimaryTextMeasurer;
            float lineWidth, lineHeight;
            textWidth = textHeight = lineWidth = lineHeight = 0;
            foreach(var r in _list)
            {
                var fontName = string.IsNullOrEmpty(r.LatinFont) ? _defaultFont.LatinFont : r.LatinFont;
                if (fontName.StartsWith("+"))
                {
                    var t=_drawing._drawings._package.Workbook.ThemeManager.GetOrCreateTheme();
                    fontName = t.GetFontByCode(fontName);
                }
                var f = new MeasurementFont()
                {
                    FontFamily = fontName,
                    Size = r.Size <= 0 ? _defaultFont.Size : r.Size,
                    Style = GetFontStyle(r)
                };
                var b = tm.MeasureText(r.Text, f);
                if (r.IsFirstInParagraph)
                {
                    lineWidth = b.Width;
                    lineHeight = b.Height;
                }
                else
                {
                    lineWidth += b.Width;
                    if (lineHeight < b.Height)
                    {
                        lineHeight = b.Height;
                    }
                }

                if (r.IsLastInParagraph)
                {
                    if(lineWidth > textWidth)
                    {
                        textWidth = lineWidth;
                    }
                    textHeight += lineHeight;
                }
            }
        }

        private MeasurementFontStyles GetFontStyle(ExcelParagraph r)
        {
            MeasurementFontStyles ret = MeasurementFontStyles.Regular;
            if (r.Bold)
            {
                ret |= MeasurementFontStyles.Bold;
            }
            if (r.Italic)
            {
                ret |= MeasurementFontStyles.Italic;
            }
            if(r.UnderLine!=eUnderLineType.None)
            {
                ret |= MeasurementFontStyles.Underline;
            }
            return ret;
        }
    }
}
