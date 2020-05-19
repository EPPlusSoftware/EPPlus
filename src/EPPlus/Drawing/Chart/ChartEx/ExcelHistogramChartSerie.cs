/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelHistogramChartSerie : ExcelChartExSerie
    {
        public ExcelHistogramChartSerie(ExcelChartEx chart, XmlNamespaceManager ns, XmlNode node) : base(chart, ns, node)
        {
            if (chart.ChartType == eChartType.Pareto)
            {
                AddParetoLine();
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
        internal void AddParetoLineFromSerie(XmlElement serElement)
        {
            ParetoLine = new ExcelChartExParetoLine(_chart, NameSpaceManager, serElement);
        }
        public void RemoveParetoLine()
        {
            ParetoLine?.DeleteNode(".");
        }
        public ExcelChartExParetoLine ParetoLine
        {
            get;
            private set;
        } = null;
    }
}
