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
            //TODO: Arguably this is just another series with a series cache.
            //Same as Cat or Val except that it is added in Ext on the Serie node
            //ShowValue property essentially changes the datalabels in the same way.
            //This could be unified somehow so that all serie ranges; Cat, Val and DataLabelRange are handled the same way.

            bool moreThanOneRow = address.Rows > 1;
            bool moreThanOneColumn = address.Columns > 1;

            if (moreThanOneRow && moreThanOneColumn)
            {
                throw new InvalidExpressionException($"DataLabelRange cannot be set to invalid range: '{address.Address}'\n" +
                    $"The range must be a single cell, a single row or a single column");
            }

            DataLabelRange = address;

            //TODO: The way we aquire the Series instance here is obtuse.
            //Fix as part of datalabel refactor?
            //Perhaps the series of a series label should be part of its constructor.
            //Or use an eventhandler
            //For a single case however that feels overkill.

            //Has to get the series index:
            var idxNode = (XmlElement)TopNode.ParentNode.SelectSingleNode($"{NsPrefix}:idx", NameSpaceManager);
            var idxNodeValue = int.Parse(idxNode.GetAttribute("val"));
            //Get the series this datalabel is on
            var currentSeries = (ExcelChartStandardSerie)_chart.Series[idxNodeValue];
            //Set the ext data needed in the Series node
            currentSeries.SetDataLabelRange(address);

            //Create the Datalabels if they do not exist
            if (DataLabels.Count < currentSeries.NumberOfItems)
            {
                for (int i = 0; i < currentSeries.NumberOfItems; i++)
                {
                    ExcelChartDataLabelItem currentLabel;
                    if (DataLabels.Count - 1 < i)
                    {
                        currentLabel = DataLabels.Add(i);
                    }
                    else
                    {
                        currentLabel = DataLabels[i];
                    }

                    //Adds field CellRange to the paragraph of the label
                    currentLabel.AddField("CELLRANGE");

                    currentLabel.AddEmptyExtFieldTableNode();
                    currentLabel.ShowDatalabelsRange = true;
                }
            }
        }

        void AddCellRangeFieldToLabel(int labelIdx)
        {

            //var fieldGuid = Guid.NewGuid();

            //DataLabels[labelIdx].OverWriteText()
            //var fieldXml = $"<a:fld id=\"{{{fieldGuid}}}\" type=\"CELLRANGE\">\r\n<a:rPr lang=\"en-US\"/>\r\n<a:pPr/>\r\n<a:t>[CELLRANGE]</a:t>\r\n</a:fld>";

        }
    }
}