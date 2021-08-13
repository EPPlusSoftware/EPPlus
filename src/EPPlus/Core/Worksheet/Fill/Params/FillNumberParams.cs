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
    public class FillNumberParams
    {
        /// <summary>
        /// The start value. If null, the first value in the row/column is used. 
        /// <seealso cref="Direction"/>
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
        /// The direction of the fill
        /// </summary>
        public eFillDirection Direction { get; set; } = eFillDirection.Column;
        /// <summary>
        /// The calculation method to use 
        /// </summary>
        public eCalculationMethod CalculationMethod { get; set; } = eCalculationMethod.Add;
        /// <summary>
        /// The number format to be appled to the range.
        /// </summary>
        public string NumberFormat { get; set; } = null;
    }
}