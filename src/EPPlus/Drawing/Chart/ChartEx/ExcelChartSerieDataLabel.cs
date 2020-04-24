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
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// Datalabel properties
    /// </summary>
    public class ExcelChartExSerieDataLabel : ExcelChartExDataLabel
    {
        internal ExcelChartExSerieDataLabel(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, string[] schemaNodeOrder)
             : base(chart, ns, node)
        {
            SchemaNodeOrder = schemaNodeOrder;
            Position = eLabelPosition.Center;
        }
        ExcelChartExDataLabelCollection _dataLabels = null;
        /// <summary>
        /// Individually formatted datalabels.
        /// </summary>
        public ExcelChartExDataLabelCollection DataLabels
        {
            get
            {
                if (_dataLabels == null)
                {
                    _dataLabels = new ExcelChartExDataLabelCollection(_chart, NameSpaceManager, TopNode, SchemaNodeOrder);
                }
                return _dataLabels;
            }
        }

    }
}
