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
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Datalabel properties
    /// </summary>
    public sealed class ExcelChartSerieDataLabel : ExcelChartDataLabelStandard
    {
       internal ExcelChartSerieDataLabel(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, string[] schemaNodeOrder)
           : base(chart, ns, node, "dLbls", schemaNodeOrder)
        {
            Position = eLabelPosition.Center;
        }
        ExcelChartDataLabelCollection _dataLabels = null;
        /// <summary>
        /// Individually formatted datalabels.
        /// </summary>
        public ExcelChartDataLabelCollection DataLabels
        {
            get
            {
                if (_dataLabels == null)
                {
                    _dataLabels = new ExcelChartDataLabelCollection(_chart, NameSpaceManager, TopNode, SchemaNodeOrder, this as ExcelChartDataLabelStandard);
                }
                return _dataLabels;
            }
        }

        /// <summary>
        /// Does the datalabels of this chart contain
        /// Value From Cells
        /// </summary>
        public bool ValueFromCells { get { return DataLabelRange != null; } }

        ExcelRangeBase DataLabelRange = null;

        /// <summary>
        /// Select datalabel range for
        /// Value From Cells
        /// </summary>
        /// <param name="address">must be a single; cell, row or column</param>
        /// <exception cref="InvalidExpressionException">Thrown when input is not a cell, a row or a column</exception>
        public void SelectRange(ExcelRangeBase address)
        {
            bool moreThanOneRow = address.Rows > 1;
            bool moreThanOneColumn = address.Columns > 1;

            if (moreThanOneRow && moreThanOneColumn)
            {
                throw new InvalidExpressionException($"DataLabelRange cannot be set to invalid range: '{address.Address}'\n" +
                    $"The range must be a single cell, a single row or a single column");
            }

            DataLabelRange = address;

            //Has to get the series index:
            var idxNode = (XmlElement)TopNode.ParentNode.SelectSingleNode($"{NsPrefix}:idx", NameSpaceManager);
            var idxNodeValue = int.Parse(idxNode.GetAttribute("val"));

            var currentSeries = (ExcelChartStandardSerie)_chart.Series[idxNodeValue];

            currentSeries.NameSpaceManager.AddNamespace("c15", ExcelPackage.schemaChart2012);
            currentSeries.NameSpaceManager.AddNamespace("c16", ExcelPackage.schemaChart2014);

            string extPath = "c:extLst/c:ext";

            XmlElement ext15Node;

            var c15Uri = "{02D57815-91ED-43cb-92C2-25804820EDAC}";

            if (currentSeries.ExistsNode(extPath+ $"[@uri='{c15Uri}']") == false)
            {
                XmlElement el = (XmlElement)currentSeries.CreateNode($"{extPath}");
                el.SetAttribute("xmlns:c15", ExcelPackage.schemaChart2012);
                currentSeries.SetXmlNodeString($"{extPath}/@uri", $"{c15Uri}");
                ext15Node = el;
            }
            else
            {
                ext15Node = (XmlElement)currentSeries.GetNode($"{extPath}");
            }

            if (currentSeries.ExistsNode($"{extPath}[2]") == false)
            {
                XmlElement element = (XmlElement)currentSeries.CreateNode($"{extPath}", false, true);
                element.SetAttribute("xmlns:c16", ExcelPackage.schemaChart2014);
                currentSeries.SetXmlNodeString($"{extPath}[2]/@uri", "{C3380CC4-5D6E-409C-BE32-E72D297353CC}");
                var _guidId = Guid.NewGuid();

                var extNode2 = currentSeries.GetNode($"{extPath}[2]");
                var uniqueIdNode = (XmlElement)CreateNode(extNode2, "c16:uniqueID");
                uniqueIdNode.SetAttribute("val", $"{{{_guidId}}}");
            }

            var dlblRangePath = $"{extPath}/c15:datalabelsRange";
            var datalabelsRange = currentSeries.CreateNode(dlblRangePath);
            var formulaNode = currentSeries.CreateNode($"{dlblRangePath}/c15:f");
            formulaNode.InnerText = address.AddressAbsolute;

            if(DataLabels.Count == 0)
            {
                for (int i = 0; i < currentSeries.NumberOfItems; i++)
                {
                    var individualLabel = DataLabels.Add(i);
                    individualLabel.AddExtFieldTableEmpty();
                    individualLabel.ShowDatalabelsRange = true;
                }
            }

            var rangeNode = currentSeries.CreateNode($"{dlblRangePath}/c15:dlblRangeCache");
            currentSeries.CreateCache(address.FullAddressAbsolute, rangeNode);
        }
    }
}