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
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A collection of trendlines.
    /// </summary>
    public class ExcelChartTrendlineCollection : IEnumerable<ExcelChartTrendline>
    {
        List<ExcelChartTrendline> _list = new List<ExcelChartTrendline>();
        ExcelChartStandardSerie _serie;
        internal ExcelChartTrendlineCollection(ExcelChartStandardSerie serie)
        {
            _serie = serie;
            if (serie != null)
            {
                foreach (XmlNode node in _serie.TopNode.SelectNodes("c:trendline", _serie.NameSpaceManager))
                {
                    _list.Add(new ExcelChartTrendline(_serie.NameSpaceManager, node, serie));
                }
            }
        }
        /// <summary>
        /// Add a new trendline
        /// </summary>
        /// <param name="Type"></param>
        /// <returns>The trendline</returns>
        public ExcelChartTrendline Add(eTrendLine Type)
        {
            if (_serie==null ||
                _serie._chart.IsType3D() ||
                _serie._chart.IsTypePercentStacked() ||    
                _serie._chart.IsTypeStacked() ||
                _serie._chart.IsTypePieDoughnut())
            {
                throw(new ArgumentException("Type","Trendlines don't apply to 3d-charts, stacked charts, pie charts, doughnut charts or Excel 2016 chart types"));
            }
            ExcelChartTrendline tl;
            XmlNode insertAfter;
            if (_list.Count > 0)
            {
                insertAfter = _list[_list.Count - 1].TopNode;
            }
            else
            {
                insertAfter = _serie.TopNode.SelectSingleNode("c:marker", _serie.NameSpaceManager);
                if (insertAfter == null)
                {
                    insertAfter = _serie.TopNode.SelectSingleNode("c:tx", _serie.NameSpaceManager);
                    if (insertAfter == null)
                    {
                        insertAfter = _serie.TopNode.SelectSingleNode("c:order", _serie.NameSpaceManager);
                    }
                }
            }

            var node=_serie.TopNode.OwnerDocument.CreateElement("c","trendline", ExcelPackage.schemaChart);
            _serie.TopNode.InsertAfter(node, insertAfter);
            node.InnerXml = "<c:trendlineLbl><c:numFmt sourceLinked=\"0\" formatCode=\"General\"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr anchorCtr=\"1\" anchor=\"ctr\" wrap=\"square\" vert=\"horz\" vertOverflow=\"ellipsis\" spcFirstLastPara=\"1\" rot=\"0\"/><a:lstStyle/><a:p><a:pPr><a:defRPr baseline=\"0\" kern=\"1200\" strike=\"noStrike\" u=\"none\" i=\"0\" b=\"0\" sz=\"900\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:pPr><a:endParaRPr/></a:p></c:txPr></c:trendlineLbl>";
            tl = new ExcelChartTrendline(_serie.NameSpaceManager, node, _serie);
            tl.Type = Type;
            _serie._chart.ApplyStyleOnPart(tl, _serie._chart.StyleManager?.Style?.Trendline);
            _serie._chart.ApplyStyleOnPart(tl.Label, _serie._chart.StyleManager?.Style?.TrendlineLabel);
            _list.Add(tl);
            return tl;
        }
        /// <summary>
        /// Number of items in the collection.
        /// </summary>
        public int Count 
        { 
            get
            {
                return _list.Count;
            }
        }
        /// <summary>
        /// Returns a chart trendline at the specific position.  
        /// </summary>
        /// <param name="index">The index in the collection. 0-base</param>
        /// <returns></returns>
        public ExcelChartTrendline this[int index]
        {
            get
            {
                if(index < 0 && index >= _list.Count)
                {
                    throw new IndexOutOfRangeException();
                }
                return _list[index];
            }
        }

        IEnumerator<ExcelChartTrendline> IEnumerable<ExcelChartTrendline>.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
    }
}