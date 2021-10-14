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
using System.Collections.Generic;

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
        /// Fill is done by column from top to bottom
        /// </summary>
        /// <param name="startValue">The start value of the first cell. If this value is null the value of the first cell is used.</param>
        /// <param name="stepValue">The value used for each step</param>
        public void FillNumber(double? startValue, double stepValue=1)
        {
            FillNumber(x => { x.StepValue = stepValue; x.StartValue = startValue; });
        }
        public void FillNumber(Action<FillNumberParams> options)
        {
            var o = new FillNumberParams();
            options?.Invoke(o);

            if (o.Direction == eFillDirection.Column)
            {
                if (o.StartPosition == eFillStartPosition.TopLeft)
                {
                    for (int c = _fromCol; c <= _toCol; c++)
                    {
                        FillMethods.FillNumber(_worksheet, _fromRow, _toRow, c, c, o);
                    }
                }
                else
                {
                    for (int c = _toCol; c >= _fromCol; c--)
                    {
                        FillMethods.FillNumber(_worksheet, _fromRow, _toRow, c, c, o);
                    }
                }
            }
            else
            {
                if (o.StartPosition == eFillStartPosition.TopLeft)
                {
                    for (int r = _fromRow; r <= _toRow; r++)
                    {
                        FillMethods.FillNumber(_worksheet, r, r, _fromCol, _toCol, o);
                    }
                }
                else
                {
                    for (int r = _toRow; r >= _fromRow; r--)
                    {
                        FillMethods.FillNumber(_worksheet, r, r, _fromCol, _toCol, o);
                    }
                }
            }

            if (!string.IsNullOrEmpty(o.NumberFormat))
            {
                Style.Numberformat.Format = o.NumberFormat;
            }
        }
         #endregion
        #region FillDateTime
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
        public void FillDateTime(DateTime? startValue, eDateTimeUnit dateTimeUnit=eDateTimeUnit.Day, int stepValue = 1)
        {
            FillDateTime(x => 
            {
                x.StartValue = startValue;
                x.DateTimeUnit = dateTimeUnit;
                x.StepValue = stepValue;
            });
        }
        /// <summary>
        /// Fill the range with dates.
        /// </summary>
        /// <param name="options">Options how to perform the fill</param>
        public void FillDateTime(Action<FillDateParams> options)
        {
            var o = new FillDateParams();
            options?.Invoke(o);

            if (o.Direction == eFillDirection.Column)
            {
                if (o.StartPosition == eFillStartPosition.TopLeft)
                {
                    for (int c = _fromCol; c <= _toCol; c++)
                    {
                        FillMethods.FillDateTime(_worksheet, _fromRow, _toRow, c, c, o);
                    }
                }
                else
                {
                    for (int c = _toCol; c >= _fromCol; c--)
                    {
                        FillMethods.FillDateTime(_worksheet, _fromRow, _toRow, c, c, o);
                    }
                }
            }
            else
            {
                if (o.StartPosition == eFillStartPosition.TopLeft)
                {
                    for (int r = _fromRow; r <= _toRow; r++)
                    {
                        FillMethods.FillDateTime(_worksheet, r, r, _fromCol, _toCol, o);
                    }
                }
                else
                {
                    for (int r = _toRow; r >= _fromRow; r--)
                    {
                        FillMethods.FillDateTime(_worksheet, r, r, _fromCol, _toCol, o);
                    }
                }
            }
            
            if (!string.IsNullOrEmpty(o.NumberFormat))
            {
                Style.Numberformat.Format = o.NumberFormat;
            }
        }
        #endregion
        #region FillList
        /// <summary>
        /// Fills the range columnwise using the values in the list. 
        /// </summary>
        /// <typeparam name="T">Type used in the list.</typeparam>
        /// <param name="list">The list to use.</param>
        public void FillList<T>(IEnumerable<T> list)
        {
            FillList(list, x=> { });
        }
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="options"></param>
        public void FillList<T>(IEnumerable<T> list, Action<FillListParams> options)
        {
            var o = new FillListParams();
            options?.Invoke(o);

            if (o.Direction == eFillDirection.Column)
            {
                if (o.StartPosition == eFillStartPosition.TopLeft)
                {
                    for (int c = _fromCol; c <= _toCol; c++)
                    {
                        FillMethods.FillList(_worksheet, _fromRow, _toRow, c, c, list, o);
                    }
                }
                else
                {
                    for (int c = _toCol; c >= _fromCol; c--)
                    {
                        FillMethods.FillList(_worksheet, _fromRow, _toRow, c, c, list, o);
                    }
                }
            }
            else
            {
                if (o.StartPosition == eFillStartPosition.TopLeft)
                {
                    for (int r = _fromRow; r <= _toRow; r++)
                    {
                        FillMethods.FillList(_worksheet, r, r, _fromCol, _toCol, list, o);
                    }
                }
                else
                {
                    for (int r = _toRow; r >= _fromRow; r--)
                    {
                        FillMethods.FillList(_worksheet, r, r, _fromCol, _toCol, list, o);
                    }
                }
            }

            if (!string.IsNullOrEmpty(o.NumberFormat))
            {
                Style.Numberformat.Format = o.NumberFormat;
            }
        }
        #endregion
    }
}
