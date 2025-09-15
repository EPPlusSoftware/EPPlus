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
    public class ExcelDrawingParagraphTabStopCollection : XmlHelper, IEnumerable<ExcelDrawingParagraphTabStop>
    {
        List<ExcelDrawingParagraphTabStop> _tabStops = new List<ExcelDrawingParagraphTabStop>();
        internal ExcelDrawingParagraphTabStopCollection(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, Action initXml) : base(nameSpaceManager, topNode)
        {
            var pNodes = topNode.SelectNodes("a:pPr/a:tabLst/a:tab", nameSpaceManager);

            foreach(XmlElement pn in pNodes)
            {
                _tabStops.Add(new ExcelDrawingParagraphTabStop(nameSpaceManager, pn, schemaNodeOrder,  initXml));
            }
        }
        public int Count { get => _tabStops.Count; }
        public ExcelDrawingParagraphTabStop this[int PositionID]
        {
            get
            {
                return _tabStops[PositionID];
            }
        }
        /// <summary>
        /// Gets the enumerator.
        /// </summary>
        /// <returns></returns>
        public IEnumerator<ExcelDrawingParagraphTabStop> GetEnumerator()
        {
            return _tabStops.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}