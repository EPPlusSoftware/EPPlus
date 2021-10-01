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
namespace OfficeOpenXml.Core.Worksheet.Fill
{
    /// <summary>
    /// Parameters for the <see cref="ExcelRangeBase.FillNumber(System.Action{FillNumberParams})" /> method 
    /// </summary>
    public class FillNumberParams : FillParams
    {
        /// <summary>
        /// The start value. If null, the first value in the row/column is used. 
        /// <seealso cref="FillParams.Direction"/>
        /// </summary>
        public double? StartValue { get; set; } = null;
        /// <summary>
        /// When this value is exceeded the fill stops
        /// </summary>
        public double? EndValue { get; set; } = null;
        /// <summary>
        /// The value to use in the calculation for each step. 
        /// <seealso cref="CalculationMethod"/>
        /// </summary>
        public double StepValue { get; set; } = 1;
        /// <summary>
        /// The calculation method to use 
        /// </summary>
        public eCalculationMethod CalculationMethod { get; set; } = eCalculationMethod.Add;
    }
}