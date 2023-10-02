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
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Xml;
namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// Represents an Autofilter for a worksheet or a filter of a table
    /// </summary>
    public class ExcelAutoFilter : XmlHelper
    {
        private const string AutoFilterGuid= "71E0E44A-7884-43F4-9E11-E314B2584A5E";
        private ExcelWorksheet _worksheet;
        private ExcelTable _table;

        internal ExcelAutoFilter(XmlNamespaceManager namespaceManager, XmlNode topNode, ExcelWorksheet worksheet) : base(namespaceManager, topNode)
        {
            _columns = new ExcelFilterColumnCollection(namespaceManager, topNode, this);
            _worksheet = worksheet;
            if (GetXmlNodeString("d:autoFilter/@ref") != "")
            {
                Address = new ExcelAddressBase(GetXmlNodeString("d:autoFilter/@ref"));
            }
        }
        internal ExcelAutoFilter(XmlNamespaceManager namespaceManager, XmlNode topNode, ExcelTable table) : base(namespaceManager, topNode)
        {
            _columns = new ExcelFilterColumnCollection(namespaceManager, topNode, this);
            _worksheet = table.WorkSheet;
            _table = table;
            if(GetXmlNodeString("d:autoFilter/@ref") != "")
            {
                Address = new ExcelAddressBase(GetXmlNodeString("d:autoFilter/@ref"));
            }
        }

        internal void Save()
        {
            ApplyFilter();
            foreach (var c in Columns)
            {
                c.Save();
            }
        }
        /// <summary>
        /// Applies the filter, hiding rows not matching the filter columns
        /// </summary>
        /// <param name="calculateRange">If true, any formula in the autofilter range will be calculated before the filter is applied.</param>
        public void ApplyFilter(bool calculateRange=false)
        {
            if(calculateRange && _address!=null && ExcelAddressBase.IsValidAddress(_address._address))
            {
                _worksheet.Cells[_address._address].Calculate();
            }

            foreach (var column in Columns)
            {
                column.SetFilterValue(_worksheet, Address);
            }
            for (int row=Address._fromRow+1; row <= _address._toRow;row++)
            {
                var rowInternal = ExcelRow.GetRowInternal(_worksheet, row);
                rowInternal.Hidden = false;
                foreach(var column in Columns)
                {
                    var value = _worksheet.GetCoreValueInner(row, Address._fromCol + column.Position);
                    var text = ValueToTextHandler.GetFormattedText(value._value, _worksheet.Workbook, value._styleId, false);
                    if (column.Match(value._value, text) == false)
                    {
                        rowInternal.Hidden = true;
                        break;
                    }
                }
            }
        }

        ExcelAddressBase _address = null;

        /// <summary>
        /// The range of the autofilter
        /// Autofilter with address "" or null indicates empty autofilter.
        /// </summary>
        public ExcelAddressBase Address
        {
            get
            {
                _worksheet.CheckSheetTypeAndNotDisposed();
                if(_address == null)
                {
                    string address = GetXmlNodeString("d:autoFilter/@ref");
                    if (address == "")
                    {
                        _address = null;
                    }
                    else
                    {
                        if(_address.Address != address)
                        {
                            _address = new ExcelAddressBase(address);
                        }
                    }
                }

                return _address;
            }
            internal set
            {
                _worksheet.CheckSheetTypeAndNotDisposed();

                if (value == null)
                {
                    DeleteAllNode($"d:autoFilter/@ref");
                    _columns = null;
                }
                else
                {
                    if (_address == null || value._fromCol != Address._fromCol || value._toCol != Address._toCol || value._fromRow != Address._fromRow) //Allow different _toRow
                    {
                        _columns = new ExcelFilterColumnCollection(NameSpaceManager, TopNode, this);
                    }

                    SetXmlNodeString($"d:autoFilter/@ref", value.Address);
                }

                _address = value;
            }
        }

        ExcelFilterColumnCollection _columns;
        /// <summary>
        /// The columns to filter
        /// </summary>
        public ExcelFilterColumnCollection Columns
        {
            get
            {
                return _columns;
            }
        }

        /// <summary>
        /// Clear all columns Unhide all affected cells, nullify address and table.
        /// </summary>
        public void ClearAll()
        {
            _worksheet.Cells[_address.Address].EntireRow.Hidden = false;
            _columns.Clear();
            _table = null;
            Address = null;
        }
        internal XmlNode CreateAutoFilterTopNode()
        {
            if(_table==null)
            {
                return _worksheet.CreateNode("d:autoFilter");
            }
            else
            {
                return _table.CreateNode("d:autoFilter");
            }
        }
    }
}
