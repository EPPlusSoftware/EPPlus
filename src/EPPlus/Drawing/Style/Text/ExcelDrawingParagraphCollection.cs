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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    public class ExcelDrawingParagraphCollection : XmlHelper, IEnumerable<ExcelDrawingParagraph>
    {
        IPictureRelationDocument _prd;
        string _path;
        Action _initXml = null;
        List<ExcelDrawingParagraph> _paragraphs = new List<ExcelDrawingParagraph>();
        internal ExcelDrawingParagraphCollection(IPictureRelationDocument prd, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string path, string[] schemaNodeOrder, Action initXml) : base(nameSpaceManager, topNode)
        {
            _prd = prd;
            var pNodes = topNode.SelectNodes(path + "/a:p", nameSpaceManager);
            _path = path;
            foreach (XmlElement pn in pNodes)
            {
                _paragraphs.Add(new ExcelDrawingParagraph(prd, nameSpaceManager, pn, schemaNodeOrder, initXml));
            }
        }
        public int Count { get => _paragraphs.Count; }

        public bool IsReadOnly => throw new NotImplementedException();

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
            if(_paragraphs.Count==0)
            {
                CreateTopNode();
            }
            var pn = CreateNode("a:p", false, true);
            var p = new ExcelDrawingParagraph(_prd, NameSpaceManager, pn, SchemaNodeOrder, _initXml);
            p.TextRuns.Add(text);
            _paragraphs.Add(p);
            return p;
        }
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= _paragraphs.Count)
            {
                throw new IndexOutOfRangeException("Paragraph index out of range.");
            }
            var pn = _paragraphs[index].TopNode;
            pn.ParentNode.RemoveChild(pn);
            _paragraphs.RemoveAt(index);
        }
        public void Remove(ExcelDrawingParagraph item)
        {
            if (!_paragraphs.Contains(item))
            {
                throw new ArgumentException("Paragraph item does not exist in the collection");
            }
            var pn = item.TopNode;
            pn.ParentNode.RemoveChild(pn);
            _paragraphs.Remove(item);
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
        /// Clears all para
        /// </summary>
        public void Clear()
        {
            for (int i = 0; i < _paragraphs.Count; i++)
            {
                var pn= _paragraphs[i].TopNode;
                pn.ParentNode.RemoveChild(pn);
            }
            _paragraphs.Clear();
        }

        public bool Contains(ExcelDrawingParagraph item)
        {
            return _paragraphs.Contains(item);
        }
        /// <summary> 
        /// Creates the top nodes of the collection
        /// </summary>
        protected internal void CreateTopNode()
        {
            if (_paragraphs.Count == 0)
            {
                _initXml?.Invoke();
                TopNode = CreateNode(_path);
                CreateNode("a:bodyPr");
                CreateNode("a:lstStyle");
            }
        }
    }

}