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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;
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
        /// <param name="insertType">Wheter Inserting rows or columns</param>
        internal void UpdateAddressesInsert(int fromCol, int columns, ExcelAddressBase affectedRange, eShiftTypeInsert insertType)
        {
            Series = UpdateAddressString(Series, fromCol, columns, affectedRange, insertType);
            XSeries = UpdateAddressString(XSeries, fromCol, columns, affectedRange, insertType);

            if(HeaderAddress != null && string.IsNullOrEmpty(HeaderAddress.FullAddress) == false)
            {
                var hAddress = UpdateAddressString(HeaderAddress.FullAddress, fromCol, columns, affectedRange, insertType);
                HeaderAddress = new ExcelAddressBase(hAddress);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fromCol"></param>
        /// <param name="columns"></param>
        /// <param name="affectedRange"></param>
        /// <param name="insertType">Wheter Inserting rows or columns</param>
        internal void UpdateAddressesDelete(int fromCol, int columns, ExcelAddressBase affectedRange, eShiftTypeDelete insertType)
        {
            Series = UpdateAddressStringDelete(Series, fromCol, columns, affectedRange, insertType);
            XSeries = UpdateAddressStringDelete(XSeries, fromCol, columns, affectedRange, insertType);

            if (HeaderAddress != null && string.IsNullOrEmpty(HeaderAddress.FullAddress) == false)
            {
                var hAddress = UpdateAddressStringDelete(HeaderAddress.FullAddress, fromCol, columns, affectedRange, insertType);
                HeaderAddress = new ExcelAddressBase(hAddress);
            }
        }

        string UpdateAddressStringDelete(string address, int from, int numToDelete, ExcelAddressBase affectedRange, eShiftTypeDelete insertType)
        {
            if (string.IsNullOrEmpty(address) == false)
            {
                if (ExcelCellBase.IsValidAddress(address))
                {
                    var addressBase = new ExcelAddressBase(address);
                    if (address != null && affectedRange.Collide(addressBase) != ExcelAddressBase.eAddressCollition.No)
                    {
                        if (insertType == eShiftTypeDelete.Left)
                        {
                            addressBase = addressBase.DeleteColumn(from, numToDelete);
                        }
                        else if (insertType == eShiftTypeDelete.Up)
                        {
                            addressBase = addressBase.DeleteRow(from, numToDelete);
                        }
                    }

                    if(addressBase == null)
                    {
                        return "";
                    }
                    else
                    {
                        return addressBase.FullAddress;
                    }
                }
            }
            return address;
        }


        string UpdateAddressString(string address, int from, int numToAdd, ExcelAddressBase affectedRange, eShiftTypeInsert insertType)
        {
            if(string.IsNullOrEmpty(address) == false)
            {
                if (ExcelCellBase.IsValidAddress(address))
                {
                    var addressBase = new ExcelAddressBase(address);
                    if (address != null && affectedRange.Collide(addressBase) != ExcelAddressBase.eAddressCollition.No)
                    {
                        if(insertType == eShiftTypeInsert.Right)
                        {
                            addressBase = addressBase.AddColumn(from, numToAdd);
                        }
                        else if (insertType == eShiftTypeInsert.Down)
                        {
                            addressBase = addressBase.AddRow(from, numToAdd);
                        }
                    }
                    return addressBase.FullAddress;
                }
            }
            return address;
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

        internal string GetHeaderText(int index)
        {
            var ret = "";
            if(Header != null)
            {
                return Header;
            }
            else if (HeaderAddress != null)
            {
                return GetAddressValue(HeaderAddress);
            }

            return $"Series{index + 1}";
        }

        private string GetAddressValue(ExcelAddressBase address)
        {
            if (address.IsExternal)
            {
                var wb = _chart.WorkSheet.Workbook;
                if (wb.ExternalLinks.Count < address.ExternalReferenceIndex) return ExcelErrorValue.Values.Ref;
                var extWb = wb.ExternalLinks[address.ExternalReferenceIndex - 1] as ExcelExternalWorkbook;
                if(extWb!=null)
                {
                    if (extWb.Package == null)
                    {
                        var ws = extWb.CachedWorksheets[address.WorkSheetName];
                        return ws.CellValues[address._fromRow, address._fromCol].Value.ToString();
                    }
                    else
                    {
                        var ws = extWb.Package.Workbook.Worksheets[HeaderAddress.WorkSheetName];
                        if (ws != null)
                        {
                            return ws.Cells[HeaderAddress.Address].Offset(0, 0).Text;
                        }
                    }
                }
            }
            else
            {
                ExcelWorksheet ws;
                if (string.IsNullOrEmpty(HeaderAddress.WorkSheetName))
                {
                    ws = _chart.WorkSheet;
                }
                else
                {
                    ws = _chart.WorkSheet.Workbook.Worksheets[HeaderAddress.WorkSheetName];
                }
                if (ws != null)
                {
                    if (HeaderAddress.IsSingleCell)
                    {
                        return ws.Cells[HeaderAddress.Address].Offset(0, 0).Text;
                    }
                    else
                    {
                        var sb = new StringBuilder();
                        foreach (var cell in ws.Cells[HeaderAddress.Address])
                        {
                            if (sb.Length != 0)
                            {
                                sb.Append(" ");
                            }
                            sb.Append(cell.TextMerged);
                        }
                        return sb.ToString();
                    }
                }
            }
            return ExcelErrorValue.Values.Ref;
        }
    }
}
