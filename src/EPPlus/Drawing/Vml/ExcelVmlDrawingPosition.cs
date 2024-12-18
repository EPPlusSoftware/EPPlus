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
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// The position of a VML drawing. Used for comments and Digital Signature Lines
    /// </summary>
    public class ExcelVmlDrawingPosition : ExcelPositionBase
    {
        int _startPos;
        //int _column, _row, _columnOff, _rowOff;

        const string anchorFormat = "{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}";

        internal ExcelVmlDrawingPosition(XmlNamespaceManager ns, XmlNode topNode, int startPos) : 
            base(ns, topNode, null)
        {
            _startPos = startPos;
            Load();
            ////_column = GetXmlNodeInt(colPath);
            //string anchor = GetXmlNodeString("x:Anchor");
            //string[] numbers = anchor.Split(',');
        }

        //Hide slightly different worded otherwise public props
        internal new int RowOff
        {
            get { return RowOffset; }
            set { RowOffset = value; }
        }
        internal new int ColumnOff
        {
            get { return ColumnOffset; }
            set { ColumnOffset = value; }
        }

        /// <summary>
        /// Row offset in pixels. Zero based
        /// Row Offset in EMU
        /// 
        /// EMU units   1cm         =   1/360000 
        ///             1US inch    =   1/914400
        ///             1pixel      =   1/9525
        /// </summary>
        public int RowOffset
        {
            get
            {
                //return GetNumber(3);
                return _rowOff;
            }
            set
            {
                //SetNumber(3, value);
                _rowOff = value;
            }
        }


        /// <summary>
        /// Column offset. Zero based
        /// Column Offset in EMU
        /// 
        /// EMU units   1cm         =   1/360000 
        ///             1US inch    =   1/914400
        ///             1pixel      =   1/9525
        /// </summary>
        public int ColumnOffset
        {
            get
            {
                //return GetNumber(1);
                return _columnOff;
            }
            set
            {
                //SetNumber(1, value);
                _columnOff = value;
            }
        }

        //private void SetNumber(int pos, int value)
        //{
        //    string anchor = GetXmlNodeString("x:Anchor");
        //    string[] numbers = anchor.Split(',');

        //    if (numbers.Length == 8)
        //    {
        //        numbers[_startPos + pos] = value.ToString();
        //    }
        //    else
        //    {
        //        var size = numbers.Length;
        //        Array.Resize<string>(ref numbers, 8);
        //        for (int i = 0; i < 8; i++)
        //        {
        //            if(string.IsNullOrEmpty(numbers[i]))
        //            {
        //                numbers[i] = "0";
        //            }
        //        }
        //    }

        //    var outString = string.Format(anchorFormat, numbers[0], numbers[1], numbers[2], numbers[3], numbers[4], numbers[5], numbers[6], numbers[7]);
        //    SetXmlNodeString("x:Anchor", outString);
        //}
        private void SetNumbers()
        {
            string anchor = GetXmlNodeString("x:Anchor");
            string[] numbers = anchor.Split(',');

            if (numbers.Length != 8)
            {
                var size = numbers.Length;
                Array.Resize<string>(ref numbers, 8);
                for (int i = 0; i < 8; i++)
                {
                    if (string.IsNullOrEmpty(numbers[i]))
                    {
                        numbers[i] = "0";
                    }
                }
            }

            numbers[_startPos] = Column.ToString();
            numbers[_startPos + 1] = ColumnOffset.ToString();
            numbers[_startPos + 2] = Row.ToString();
            numbers[_startPos + 3] = RowOffset.ToString();

            var outString = string.Format(anchorFormat, numbers[0], numbers[1], numbers[2], numbers[3], numbers[4], numbers[5], numbers[6], numbers[7]);
            SetXmlNodeString("x:Anchor", outString);
        }

        private int GetNumber(int pos)
        {
            string anchor = GetXmlNodeString("x:Anchor");
            string[] numbers = anchor.Split(',');
            if (numbers.Length == 8)
            {
                int ret;
                if (int.TryParse(numbers[_startPos + pos], NumberStyles.Number, CultureInfo.InvariantCulture, out ret))
                {
                    return ret;
                }
            }
            return 0;
        }
        /// <summary>
        /// Load xml data
        /// </summary>
        public override void Load()
        {
            _column = GetNumber(0);
            _columnOff = GetNumber(1);
            _row = GetNumber(2);
            _rowOff = GetNumber(3);
        }
        /// <summary>
        /// Update xml data
        /// </summary>
        public override void UpdateXml()
        {
            SetNumbers();
            //SetXmlNodeString(colPath, _column.ToString());
            //SetXmlNodeString(colOffPath, _columnOff.ToString());
            //SetXmlNodeString(rowPath, _row.ToString());
            //SetXmlNodeString(rowOffPath, _rowOff.ToString());
        }
    }
}
