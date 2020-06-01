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
        internal ExcelChartExSerieDataLabel(ExcelChartExSerie serie, XmlNamespaceManager ns, XmlNode node, string[] schemaNodeOrder)
             : base(serie, ns, node)
        {
            SchemaNodeOrder = schemaNodeOrder;
            Position = eLabelPosition.Center;
        }
        ExcelChartExDataLabelCollection _dataLabels = null;
        /// <summary>
        /// Individually formatted data labels.
        /// </summary>
        public ExcelChartExDataLabelCollection DataLabels
        {
            get
            {
                if (_dataLabels == null)
                {
                    _dataLabels = new ExcelChartExDataLabelCollection(_serie, NameSpaceManager, TopNode, SchemaNodeOrder);
                }
                return _dataLabels;
            }
        }
        /// <summary>
        /// Adds data labels to the series.
        /// </summary>
        /// <param name="showCategory">Show the category name</param>
        /// <param name="showValue">Show values</param>
        /// <param name="showSeriesName">Show series name</param>
        public void Add(bool showCategory=true, bool showValue=false, bool showSeriesName=false)
        {
            SetDataLabelNode();
            ShowCategory = showCategory;
            ShowValue = showValue;
            ShowSeriesName = showSeriesName;
        }
        /// <summary>
        /// Removes data labels from the series
        /// </summary>
        public void Remove()
        {
            _serie.DeleteNode("cx:dataLabels");
        }        
    }
}
