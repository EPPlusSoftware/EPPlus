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
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Toplevel class for cell styling
    /// </summary>
    public sealed class ExcelStyle : StyleBase
    {
        internal ExcelStyle(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int positionID, string Address, int xfsId) :
            base(styles, ChangedEvent, positionID, Address)
        {
            Index = xfsId;
            ExcelXfs xfs;
            if (positionID > -1)
            {
                xfs = _styles.CellXfs[xfsId];
            }
            else
            {
                if (_styles.CellStyleXfs.Count == 0)   //CellStyleXfs.Count should never be 0, but for some custom build sheets this can happend.
                {
                    var item=_styles.CellXfs[0].Copy();                    
                    _styles.CellStyleXfs.Add(item.Id, item);
                }
                xfs = _styles.CellStyleXfs[xfsId];
            }
            Styles = styles;
            PositionID = positionID;
            Numberformat = new ExcelNumberFormat(styles, ChangedEvent, PositionID, Address, xfs.NumberFormatId);
            Font = new ExcelFont(styles, ChangedEvent, PositionID, Address, xfs.FontId);
            Fill = new ExcelFill(styles, ChangedEvent, PositionID, Address, xfs.FillId);
            Border = new Border(styles, ChangedEvent, PositionID, Address, xfs.BorderId); 
        }
        /// <summary>
        /// Numberformat
        /// </summary>
        public ExcelNumberFormat Numberformat { get; set; }
        /// <summary>
        /// Font styling
        /// </summary>
        public ExcelFont Font { get; set; }
        /// <summary>
        /// Fill Styling
        /// </summary>
        public ExcelFill Fill { get; set; }
        /// <summary>
        /// Border 
        /// </summary>
        public Border Border { get; set; }
        /// <summary>
        /// The horizontal alignment in the cell
        /// </summary>
        public ExcelHorizontalAlignment HorizontalAlignment
        {
            get
            {
                return _styles.CellXfs[Index].HorizontalAlignment;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.HorizontalAlign, value, _positionID, _address));
            }
        }
        /// <summary>
        /// The vertical alignment in the cell
        /// </summary>
        public ExcelVerticalAlignment VerticalAlignment
        {
            get
            {
                return _styles.CellXfs[Index].VerticalAlignment;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.VerticalAlign, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Wrap the text
        /// </summary>
        public bool WrapText
        {
            get
            {
                return _styles.CellXfs[Index].WrapText;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.WrapText, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Readingorder
        /// </summary>
        public ExcelReadingOrder ReadingOrder
        {
            get
            {
                return _styles.CellXfs[Index].ReadingOrder;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.ReadingOrder, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Makes the text vertically. This is the same as setting <see cref="TextRotation"/> to 255.
        /// </summary>
        public void SetTextVertical()
        {
            TextRotation = 255;
        }

        /// <summary>
        /// Shrink the text to fit
        /// </summary>
        public bool ShrinkToFit
        {
            get
            {
                return _styles.CellXfs[Index].ShrinkToFit;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.ShrinkToFit, value, _positionID, _address));
            }
        }
        /// <summary>
        /// The margin between the border and the text
        /// </summary>
        public int Indent
        {
            get
            {
                return _styles.CellXfs[Index].Indent;
            }
            set
            {
                if (value <0 || value > 250)
                {
                    throw(new ArgumentOutOfRangeException("Indent must be between 0 and 250"));
                }
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.Indent, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Text orientation in degrees. Values range from 0 to 180 or 255. 
        /// Setting the rotation to 255 will align text vertically.
        /// </summary>
        public int TextRotation
        {
            get
            {
                return _styles.CellXfs[Index].TextRotation;
            }
            set
            {
                if ((value < 0 || value > 180) && value!=255)
                {
                    throw new ArgumentOutOfRangeException("TextRotation out of range.");
                }
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.TextRotation, value, _positionID, _address));
            }
        }
        /// <summary>
        /// If true the cell is locked for editing when the sheet is protected
        /// <seealso cref="ExcelWorksheet.Protection"/>
        /// </summary>
        public bool Locked
        {
            get
            {
                return _styles.CellXfs[Index].Locked;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.Locked, value, _positionID, _address));
            }
        }
        /// <summary>
        /// If true the formula is hidden when the sheet is protected.
        /// <seealso cref="ExcelWorksheet.Protection"/>
        /// </summary>
        public bool Hidden
        {
            get
            {
                return _styles.CellXfs[Index].Hidden;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.Hidden, value, _positionID, _address));
            }
        }

        /// <summary>
        /// If true the cell has a quote prefix, which indicates the value of the cell is text.
        /// </summary>
        public bool QuotePrefix
        {
            get
            {
                return _styles.CellXfs[Index].QuotePrefix;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.QuotePrefix, value, _positionID, _address));
            }
        }

        const string xfIdPath = "@xfId";
        /// <summary>
        /// The index in the style collection
        /// </summary>
        public int XfId 
        {
            get
            {
                return _styles.CellXfs[Index].XfId;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.XfId, value, _positionID, _address));
            }
        }
        internal int PositionID
        {
            get;
            set;
        }
        internal ExcelStyles Styles
        {
            get;
            set;
        }
        internal override string Id
        {
            get 
            { 
                return Numberformat.Id + "|" + Font.Id + "|" + Fill.Id + "|" + Border.Id + "|" + VerticalAlignment + "|" + HorizontalAlignment + "|" + WrapText.ToString() + "|" + ReadingOrder.ToString() + "|" + XfId.ToString() + "|" + QuotePrefix.ToString(); 
            }
        }

    }
}
