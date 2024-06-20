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
using OfficeOpenXml.Style;
using System.Xml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Vml;

namespace OfficeOpenXml
{
    /// <summary>
    /// An Excel Cell Comment
    /// </summary>
    public class ExcelComment : ExcelVmlDrawingComment
    {
        internal XmlHelper _commentHelper;
        private string _text;
        internal ExcelComment(XmlNamespaceManager ns, XmlNode commentTopNode, ExcelRangeBase cell)
            : base(null, cell, cell.Worksheet.VmlDrawings.NameSpaceManager)
        {
            //_commentHelper = new XmlHelper(ns, commentTopNode);
            _commentHelper = XmlHelperFactory.Create(ns, commentTopNode);
            var textElem=commentTopNode.SelectSingleNode("d:text", ns);
            if (textElem == null)
            {
                textElem = commentTopNode.OwnerDocument.CreateElement("text", ExcelPackage.schemaMain);
                commentTopNode.AppendChild(textElem);
            }
            if (!cell.Worksheet._vmlDrawings.ContainsKey(cell.Start.Row, cell.Start.Column))
            {
                cell.Worksheet._vmlDrawings.AddComment(cell);
            }

            TopNode = cell.Worksheet.VmlDrawings[cell.Start.Row, cell.Start.Column].TopNode;
            RichText = new ExcelRichTextCollection(ns, textElem, cell);
            var tNode = textElem.SelectSingleNode("d:t", ns);
            if (tNode != null)
            {
                _text = tNode.InnerText;
            }
        }
        const string AUTHORS_PATH = "d:comments/d:authors";
        const string AUTHOR_PATH = "d:comments/d:authors/d:author";
        /// <summary>
        /// The author
        /// </summary>
        public string Author
        {
            get
            {
                int authorRef = _commentHelper.GetXmlNodeInt("@authorId");
                return _commentHelper.TopNode.OwnerDocument.SelectSingleNode(string.Format("{0}[{1}]", AUTHOR_PATH, authorRef+1), _commentHelper.NameSpaceManager).InnerText;
            }
            set
            {
                int authorRef = GetAuthor(value);
                _commentHelper.SetXmlNodeString("@authorId", authorRef.ToString());
            }
        }
        private int GetAuthor(string value)
        {
            int authorRef = 0;
            bool found = false;
            foreach (XmlElement node in _commentHelper.TopNode.OwnerDocument.SelectNodes(AUTHOR_PATH, _commentHelper.NameSpaceManager))
            {
                if (node.InnerText == value)
                {
                    found = true;
                    break;
                }
                authorRef++;
            }
            if (!found)
            {
                var elem = _commentHelper.TopNode.OwnerDocument.CreateElement("d", "author", ExcelPackage.schemaMain);
                _commentHelper.TopNode.OwnerDocument.SelectSingleNode(AUTHORS_PATH, _commentHelper.NameSpaceManager).AppendChild(elem);
                elem.InnerText = value;
            }
            return authorRef;
        }
        /// <summary>
        /// The comment text 
        /// </summary>
        public string Text
        {
            get
            {
                if(!string.IsNullOrEmpty(RichText.Text)) return RichText.Text;
                return _text;
            }
            set
            {
                RichText.Text = value;
            }
        }
        /// <summary>
        /// Sets the font of the first richtext item.
        /// </summary>
        public ExcelRichText Font
        {
            get
            {
                if (RichText.Count > 0)
                {
                    return RichText[0];
                }
                return null;
            }
        }
        /// <summary>
        /// Richtext collection
        /// </summary>
        public ExcelRichTextCollection RichText 
        { 
           get; 
           set; 
        }

        /// <summary>
        /// Reference
        /// </summary>
        internal string Reference
		{
			get { return _commentHelper.GetXmlNodeString("@ref"); }
            set
            {
                var a = new ExcelAddressBase(value);
                var rows = a._fromRow - Range._fromRow;
                var cols= a._fromCol - Range._fromCol;
                Range.Address = value;
                _commentHelper.SetXmlNodeString("@ref", value);

                From.Row += rows;
                To.Row += rows;

                From.Column += cols;
                To.Column += cols;

                Row = Range._fromRow - 1;
                Column = Range._fromCol - 1;
            }
        }
	}
}
