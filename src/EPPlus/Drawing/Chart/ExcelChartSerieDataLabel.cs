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
using System.Data;
using System.Linq;
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
            var parentSeries = GetParentSeries();

            var address = parentSeries.GetDataLabelRange();
            if (string.IsNullOrEmpty(address) == false)
            {
                DataLabelRange = chart.WorkSheet.Cells[address];
            }
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
                    
                    //Fill datalabel addresses
                    if(DataLabelRange != null)
                    {
                        var address = DataLabelRange;
                        for(int i = 0; i< _dataLabels.Count(); i++)
                        {
                            if (address.Rows > address.Columns)
                            {
                                _dataLabels[i].SingleCellAddressFromSeries = address.TakeSingleCell(i, 0);
                            }
                            else
                            {
                                _dataLabels[i].SingleCellAddressFromSeries = address.TakeSingleCell(0, i);
                            }
                        }
                    }
                }
                return _dataLabels;
            }
        }

        /// <summary>
        /// Does the datalabels of this chart contain
        /// Value From Cells
        /// </summary>
        public bool ValueFromCells { get { return DataLabelRange != null; } }

        internal ExcelRangeBase DataLabelRange { get; private set; } = null;


        ExcelChartStandardSerie GetParentSeries()
        {
            //TODO: The way we aquire the Series instance here is clumsy.
            //Fix as part of datalabel refactor?
            //Perhaps the series of a series label should be part of its constructor.
            //Or use an eventhandler
            //For a single case however that feels overkill.

            //Has to get the series index:
            var idxNode = (XmlElement)TopNode.ParentNode.SelectSingleNode($"{NsPrefix}:idx", NameSpaceManager);
            var idxNodeValue = int.Parse(idxNode.GetAttribute("val"));
            //Get the series this datalabel is on
            return (ExcelChartStandardSerie)_chart.Series[idxNodeValue];
        }


        public void SetValueSource(ExcelAddressBase address)
        {
            SetValueSource(address.AddressAbsolute.ToString());
        }
        /// <summary>
        /// Select datalabel range for
        /// Value From Cells
        /// </summary>
        /// <param name="address">must be a single; cell, row or column</param>
        /// <exception cref="InvalidExpressionException">Thrown when input is not a cell, a row or a column</exception>
        public void SetValueSource(string strRef)
        {
            //TODO: Arguably this is just another series with a series cache.
            //Same as Cat or Val except that it is added in Ext on the Serie node
            //ShowValue property essentially changes the datalabels in the same way.
            //This could be unified somehow so that all serie ranges; Cat, Val and DataLabelRange are handled the same way.
            //The start of this is now being done in ChartDataSource.cs

            var currentSeries = GetParentSeries();
            //Set the ext data needed in the Series node
            currentSeries.SetDataLabelRange(strRef);

            if(currentSeries.DataLabelRangeSource.RefIsValidAddress)
            {
                DataLabelRange = _chart.WorkSheet.Cells[strRef];
            }

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
                    currentLabel.AddExtFieldTableEmpty();
                    currentLabel.ShowDatalabelsRange = true;

                    if (DataLabelRange != null)
                    {
                        if (DataLabelRange.Rows > DataLabelRange.Columns)
                        {
                            currentLabel.SingleCellAddressFromSeries = DataLabelRange.TakeSingleCell(i, 0);
                        }
                        else
                        {
                            currentLabel.SingleCellAddressFromSeries = DataLabelRange.TakeSingleCell(0, i);
                        }
                    }
                }
            }
        }
    }
}