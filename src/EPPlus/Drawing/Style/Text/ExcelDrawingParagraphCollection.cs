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
        List<ExcelDrawingParagraph> _paragraphs = new List<ExcelDrawingParagraph>();
        internal ExcelDrawingParagraphCollection(IPictureRelationDocument pictureRelationDocument, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string textBodyPath, string[] schemaNodeOrder, Action initXml) : base(nameSpaceManager, topNode)
        {
            var pNodes = topNode.SelectNodes(textBodyPath + "/a:p", nameSpaceManager);

            foreach(XmlElement pn in pNodes)
            {
                _paragraphs.Add(new ExcelDrawingParagraph(pictureRelationDocument, nameSpaceManager, pn, schemaNodeOrder,  initXml));
            }
        }
        public int Count { get => _paragraphs.Count; }
        public ExcelDrawingParagraph this[int PositionID]
        {
            get
            {
                return _paragraphs[PositionID];
            }
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
    }
}