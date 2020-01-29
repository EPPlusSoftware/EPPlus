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
        ExcelChartSerie _serie;
        internal ExcelChartTrendlineCollection(ExcelChartSerie serie)
        {
            _serie = serie;
            foreach (XmlNode node in _serie.TopNode.SelectNodes("c:trendline", _serie.NameSpaceManager))
            {
                _list.Add(new ExcelChartTrendline(_serie.NameSpaceManager, node, serie));
            }
        }
        /// <summary>
        /// Add a new trendline
        /// </summary>
        /// <param name="Type"></param>
        /// <returns>The trendline</returns>
        public ExcelChartTrendline Add(eTrendLine Type)
        {
            if (_serie._chart.IsType3D() ||
                _serie._chart.IsTypePercentStacked() ||    
                _serie._chart.IsTypeStacked() ||
                _serie._chart.IsTypePieDoughnut())
            {
                throw(new ArgumentException("Type","Trendlines don't apply to 3d-charts, stacked charts, pie charts or doughnut charts"));
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

            tl = new ExcelChartTrendline(_serie.NameSpaceManager, node, _serie);
            tl.Type = Type;
            return tl;
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