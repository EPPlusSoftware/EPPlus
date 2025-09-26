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

            if (topNode.SelectSingleNode(path, ns) == null)
            {
                if (tb.Paragraphs.Count == 0)
                {
                    var paragraphParent = path.Substring(0, path.LastIndexOf('/'));
                    var tmpTop = TopNode;
                    TopNode = TopNode.SelectNodes(paragraphParent, ns)[0];
                    var placeHolderNode = tb.Paragraphs.CreateAndReturnParagraphPlaceHolder();
                    TopNode = tmpTop;
                }
            }

            var tfXml = new ExcelTextFontXml(drawing._drawings, ns, TopNode, path + "/a:pPr/a:defRPr", schemaNodeOrder);
            //if(tb.Paragraphs.Count == 0)
            //{
            //    if(tfXml._rootNode.SelectSingleNode("//a:p", tfXml.XmlHelper.NameSpaceManager) == null)
            //    {

            //        var placeHolderNode = tb.Paragraphs.CreateAndReturnParagraphPlaceHolder();
            //        tfXml.XmlHelper.TopNode = placeHolderNode;
            //    }
            //}

            _defaultFont = tfXml;

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
            ExcelParagraph item;
            if (NewParagraph || _textBody.Paragraphs.Count==0)
            {
                _textBody.Paragraphs.Add(Text);
            }
            else
            {
                p = _textBody.Paragraphs[_textBody.Paragraphs.Count - 1];
                p.TextRuns.Add(Text);
            }

            return _list[_list.Count-1];
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
