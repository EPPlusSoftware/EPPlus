/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    08/11/2021         EPPlus Software AB       EPPlus 5.8
 *************************************************************************************************/
using OfficeOpenXml.Core.Worksheet.Fill;
using System;

namespace OfficeOpenXml
{
    public partial class ExcelRangeBase
    {
        #region FillNumbers
        /// <summary>
        /// Fills the range by adding 1 to each cell starting from the value in the top left cell by column
        /// </summary>
        public void FillNumber()
        {
            FillNumber(x => { });
        }
        /// <summary>
        /// Fills a range by adding the step value to the start Value. If <paramref name="startValue"/> is null the first value in the row/column is used.
        /// </summary>
        /// <param name="startValue">The start value of the first cell. If this value is null the value of the first cell is used.</param>
        /// <param name="stepValue">The value used for each step</param>
        /// <param name="direction">Direction of the fill</param>
        public void FillNumber(double? startValue, double stepValue=1, eFillDirection direction = eFillDirection.Column)
        {
            FillNumber(x => { x.StepValue = stepValue; x.StartValue = startValue; x.Direction = direction; });
        }
        public void FillNumber(Action<FillNumberParams> o)
        {
            var options = new FillNumberParams();
            o?.Invoke(options);

            if (options.Direction == eFillDirection.Column)
            {
                for (int c = _fromCol; c <= _toCol; c++)
                {
                    FillHandler.FillNumbers(_worksheet, _fromRow, _toRow, c, c, options);
                }
            }
            else
            {
                for (int r = _fromRow; r <= _toRow; r++)
                {
                    FillHandler.FillNumbers(_worksheet, r, r, _fromCol, _toCol, options);
                }
            }

            if (!string.IsNullOrEmpty(options.NumberFormat))
            {
                Style.Numberformat.Format = options.NumberFormat;
            }
        }
        #endregion
        /// <summary>
        /// Fills the range by adding 1 day to each cell starting from the value in the top left cell by column.
        /// </summary>
        public void FillDateTime()
        {
            FillDateTime(x => { });
        }
        /// <summary>
        /// Fills the range by adding 1 day to each cell per column starting from <paramref name="startValue"/>.
        /// </summary>
        public void FillDateTime(DateTime startValue)
        {
            FillDateTime(x => x.StartValue = startValue);
        }
        /// <summary>
        /// Fill the range with dates.
        /// </summary>
        /// <param name="o">Options how to perform the fill</param>
        public void FillDateTime(Action<FillDateParams> o)
        {
            var options = new FillDateParams();
            o?.Invoke(options);

            if (options.Direction == eFillDirection.Column)
            {
                for (int c = _fromCol; c <= _toCol; c++)
                {
                    FillHandler.FillDates(_worksheet, _fromRow, _toRow, c, c, options);
                }
            }
            else
            {
                for (int r = _fromRow; r <= _toRow; r++)
                {
                    FillHandler.FillDates(_worksheet, r, r, _fromCol, _toCol, options);
                }
            }
            
            if (!string.IsNullOrEmpty(options.NumberFormat))
            {
                Style.Numberformat.Format = options.NumberFormat;
            }
        }

        public void FillString()
        {

        }
    }
}
