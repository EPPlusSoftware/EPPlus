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
using System.Xml;
using OfficeOpenXml.Drawing.Chart.ChartEx;
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A collection of chart serie for a Histogram chart.
    /// </summary>
    public class ExcelHistogramChartSeries : ExcelChartSeries<ExcelHistogramChartSerie>
    {
        /// <summary>
        /// Adds a pareto line to the serie.
        /// </summary>
        public void AddParetoLine()
        {
            if(_chart.ChartType==eChartType.Pareto)
            {
                return;
            }
            if (_chart.Axis.Length == 2)
            {
                //Add pareto axis
                var axis2 = (XmlElement)_chart._chartXmlHelper.CreateNode("cx:plotArea/cx:axis", false, true);
                axis2.SetAttribute("id", "2");
                axis2.InnerXml = "<cx:valScaling min=\"0\" max=\"1\"/><cx:units unit=\"percentage\"/><cx:tickLabels/>";
            }
            foreach(ExcelHistogramChartSerie ser in _list)
            {
                ser.AddParetoLineFromSerie((XmlElement)ser.TopNode);                
            }
        }
        /// <summary>
        /// Removes the pareto line for the serie
        /// </summary>
        public void RemoveParetoLine()
        {
            if (_chart.ChartType == eChartType.Histogram)
            {
                return;
            }
            if (_chart.Axis.Length == 2)
            {
                if (_chart.Axis.Length == 3)
                {
                    //Remove percentage axis
                    _chart.Axis[2].TopNode.ParentNode.RemoveChild(_chart.Axis[2].TopNode);
                    ((ExcelChartEx)_chart)._exAxis = null;
                    _chart._axis = new ExcelChartAxis[] { _chart._axis[0], _chart._axis[1] };
                }
            }
        }
    }
}
