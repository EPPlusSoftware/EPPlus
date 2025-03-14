using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// Base class for ExcelPosition.
    /// </summary>
    public abstract class ExcelPositionBase : XmlHelper
    {
        internal delegate void SetWidthCallback();
        SetWidthCallback _setWidthCallback;
        internal ExcelPositionBase(XmlNamespaceManager ns, XmlNode node, SetWidthCallback setWidthCallback) :
            base(ns, node)
        {
            _setWidthCallback = setWidthCallback;
        }

        internal int _column, _row, _columnOff, _rowOff;
        /// <summary>
        /// The column
        /// </summary>
        public int Column
        {
            get
            {
                return _column;
            }
            set
            {
                _column = value;
                _setWidthCallback?.Invoke();
            }
        }
        /// <summary>
        /// The row
        /// </summary>
        public int Row
        {
            get
            {
                return _row;
            }
            set
            {
                _row = value;
                _setWidthCallback?.Invoke();
            }
        }

        /// <summary>
        /// Column Offset in EMU
        /// ss
        /// EMU units   1cm         =   1/360000 
        ///             1US inch    =   1/914400
        ///             1pixel      =   1/9525
        /// </summary>
        public int ColumnOff
        {
            get
            {
                return _columnOff;
            }
            set
            {
                _columnOff = value;
                _setWidthCallback?.Invoke();
            }
        }

        /// <summary>
        /// Row Offset in EMU
        /// 
        /// EMU units   1cm         =   1/360000 
        ///             1US inch    =   1/914400
        ///             1pixel      =   1/9525
        /// </summary>
        public int RowOff
        {
            get
            {
                return _rowOff;
            }
            set
            {
                _rowOff = value;
                _setWidthCallback?.Invoke();
            }
        }
        /// <summary>
        /// Load xml data
        /// </summary>
        public abstract void Load();

        /// <summary>
        /// Update xml data
        /// </summary>
        public abstract void UpdateXml();
    }
}
