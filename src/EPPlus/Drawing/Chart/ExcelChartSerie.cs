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
using System.Linq;
using OfficeOpenXml.Core.CellStore;
using System.Globalization;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Base class for chart series for standard charts
    /// </summary>
    public abstract class ExcelChartSerie : XmlHelper, IDrawingStyleBase
    {
        internal ExcelChart _chart;
        string _prefix;
        internal ExcelChartSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, string prefix="c")
            : base(ns, node)
        {
            _chart = chart;
            _prefix = prefix;
        }
        /// <summary>
        /// The header for the chart serie
        /// </summary>
        public abstract string Header { get; set; }
        /// <summary>
        /// Literals for the Y serie, if the literal values are numeric
        /// </summary>
        virtual public double[] NumberLiteralsY { get; protected set; } = null;
        /// <summary>
        /// Literals for the X serie, if the literal values are numeric
        /// </summary>
        virtual public double[] NumberLiteralsX { get; protected set; } = null;
        /// <summary>
        /// Literals for the X serie, if the literal values are strings
        /// </summary>
        virtual public string[] StringLiteralsX { get; protected set; } = null;
        /// <summary>
        /// Literals for the Y serie, if the literal values are strings
        /// </summary>
        virtual public string[] StringLiteralsY { get; protected set; } = null;
        void IDrawingStyleBase.CreatespPr()
        {
            CreatespPrNode();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fromCol"></param>
        /// <param name="columns"></param>
        /// <param name="affectedRange"></param>
        /// <param name="delete">Whether we are deleting columns. If not deleting we are inserting</param>
        internal void UpdateAddressesColumns(int fromCol, int columns, ExcelAddressBase affectedRange, bool delete = false)
        {
            Series = UpdateAddressString(Series, fromCol, columns, affectedRange);
            XSeries = UpdateAddressString(XSeries, fromCol, columns, affectedRange);

            if(HeaderAddress != null && string.IsNullOrEmpty(HeaderAddress.FullAddress) == false)
            {
                var hAddress = UpdateAddressString(HeaderAddress.FullAddress, fromCol, columns, affectedRange);
                HeaderAddress = new ExcelAddressBase(hAddress);
            }
        }

        void UpdateInsertColumnAddress(ref ExcelAddressBase address, int fromCol, int columns, ExcelAddressBase affectedRange)
        {
            if (address != null && affectedRange.Collide(address) != ExcelAddressBase.eAddressCollition.No)
            {
                address = address.AddColumn(fromCol, columns);
            }
        }

        string UpdateAddressString(string address, int fromCol, int columns, ExcelAddressBase affectedRange)
        {
            if(string.IsNullOrEmpty(address) == false)
            {
                if (ExcelCellBase.IsValidAddress(address))
                {
                    var addressBase = new ExcelAddressBase(address);
                    UpdateInsertColumnAddress(ref addressBase, fromCol, columns, affectedRange);
                    return addressBase.FullAddress;
                }
            }
            return address;
        }

        internal void UpdateAddressesRows(int rowFrom, int rowTo)
        {
            var SeriesAddress = new ExcelAddress(Series);
            var XSeriesAddress = new ExcelAddress(XSeries);

            SeriesAddress.Addresses.Add(XSeriesAddress);
            SeriesAddress.Addresses.Add(HeaderAddress);

            SeriesAddress.Intersect(new ExcelAddress(rowFrom, 1, rowTo, 2500));
        }

        /// <summary>
        /// The header address for the serie.
        /// </summary>
        public abstract ExcelAddressBase HeaderAddress { get; set; }
        /// <summary>
        /// The address for the vertical series.
        /// </summary>
        public abstract string Series { get; set; }
        /// <summary>
        /// The address for the horizontal series.
        /// </summary>
        public abstract string XSeries { get; set; }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Access to fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(_chart, NameSpaceManager, TopNode, $"{_prefix}:spPr", SchemaNodeOrder);
                }
                return _fill;
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// Access to border properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(_chart, NameSpaceManager, TopNode, $"{_prefix}:spPr/a:ln", SchemaNodeOrder);
                }
                return _border;
            }
        }
        ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// Effects
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (_effect == null)
                {
                    _effect = new ExcelDrawingEffectStyle(_chart, NameSpaceManager, TopNode, $"{_prefix}:spPr/a:effectLst", SchemaNodeOrder);
                }
                return _effect;
            }
        }
        ExcelDrawing3D _threeD = null;
        /// <summary>
        /// 3D properties
        /// </summary>
        public ExcelDrawing3D ThreeD
        {
            get
            {   
                if (_threeD == null)
                {
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, $"{_prefix}:spPr", SchemaNodeOrder);
                }
                return _threeD;
            }
        }
        /// <summary>
        /// Number of items in the serie.
        /// </summary>
        public abstract int NumberOfItems { get; }
        /// <summary>
        /// A collection of trend lines for the chart serie.
        /// </summary>
        public abstract ExcelChartTrendlineCollection TrendLines{ get; }
        internal abstract void SetID(string id);
        internal string ToFullAddress(string value)
        {
            if (ExcelCellBase.IsValidAddress(value))
            {
                return ExcelCellBase.GetFullAddress(_chart.WorkSheet.Name, value);
            }
            else
            {
                return value;
            }
        }
    }
}
