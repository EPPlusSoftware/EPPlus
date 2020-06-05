/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB           EPPlus 5.2
 *************************************************************************************************/
using System;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// A series for an Histogram Chart
    /// </summary>
    public class ExcelHistogramChartSerie : ExcelChartExSerie
    {
        internal int _index;
        internal ExcelHistogramChartSerie(ExcelChartEx chart, XmlNamespaceManager ns, XmlNode node, int index=-1) : base(chart, ns, node)
        {
            if(index==-1)
            {
                _index = chart.Series.Count * (chart.ChartType == eChartType.Pareto ? 2 : 1);
            }
            else
            {
                _index = index;
            }
        }
        internal void AddParetoLine()
        {
            var ix = _chart.Series.Count * 2;
            var serElement = ExcelChartExSerie.CreateSeriesElement((ExcelChartEx)_chart, eChartType.Pareto, ix+1, TopNode, true);
            serElement.SetAttribute("ownerIdx", (ix).ToString());
            serElement.InnerXml = "<cx:axisId val=\"2\"/>";
            AddParetoLineFromSerie(serElement);
        }
        ExcelChartExSerieBinning _binning = null;
        /// <summary>
        /// The data binning properties
        /// </summary>
        public ExcelChartExSerieBinning Binning
        {
            get
            {
                if (_binning == null)
                {
                    _binning = new ExcelChartExSerieBinning(NameSpaceManager, TopNode);
                }
                return _binning;
            }
        }
        internal const string _aggregationPath = "cx:layoutPr/cx:aggregation";
        internal const string _binningPath = "cx:layoutPr/cx:binning";
        /// <summary>
        /// If x-axis is per category
        /// </summary>
        public bool Aggregation
        {
            get
            {
                return ExistNode(_aggregationPath);
            }
            set
            {
                if (value)
                {
                    DeleteNode(_binningPath);
                    CreateNode(_aggregationPath);
                }
                else
                {
                    DeleteNode(_aggregationPath);
                    if(!ExistNode(_binningPath))
                    {
                        Binning.IntervalClosed = eIntervalClosed.Right;
                    }
                }
            }
        }
        internal void AddParetoLineFromSerie(XmlElement serElement)
        {
            ParetoLine = new ExcelChartExParetoLine(_chart, NameSpaceManager, serElement);
        }
        internal void RemoveParetoLine()
        {
            ParetoLine?.DeleteNode(".");
            ParetoLine = null;
        }
        public ExcelChartExParetoLine ParetoLine
        {
            get;
            private set;
        } = null;
    }
}
