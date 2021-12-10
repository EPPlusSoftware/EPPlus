/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/27/2020         EPPlus Software AB           EPPlus 5.2
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// A plotarea for an extended chart
    /// </summary>
    public sealed class ExcelChartExPlotarea : ExcelChartPlotArea
    {
        public ExcelChartExPlotarea(XmlNamespaceManager ns, XmlNode node, ExcelChart chart) : base(ns, node, chart, "cx")
        {
            SchemaNodeOrder = new string[] { "plotAreaRegion","axis","spPr" };
        }
        public override ExcelChartDataTable CreateDataTable()
        {
            throw (new InvalidOperationException("Extensions charts cannot have a data tables"));
        }
        public override void RemoveDataTable()
        {
            throw (new InvalidOperationException("Extensions charts cannot have a data tables"));
        }
    }
}
