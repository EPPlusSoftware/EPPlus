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
namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    /// <summary>
    /// Result type
    /// </summary>
    public enum CompileResultType
    {
        /// <summary>
        /// A normal compile result containing a value.
        /// </summary>
        Normal = 0,
        /// <summary>
        /// A compile result referencing a range address. This will allow the result to be used with the colon operator.
        /// </summary>
        RangeAddress = 1,
        /// <summary>
        /// The result is a dynamic array formula.
        /// </summary>
        DynamicArray = 2,
        /// <summary>
        /// The result is a dynamic array formula. Even if the result is nested in another function that the cell should be marked as dynamic.
        /// </summary>
        DynamicArray_AlwaysSetCellAsDynamic = 3,
        /// <summary>
        /// A compile result containing a local image or a reference to a local image
        /// </summary>
        LocalImage = 4,
        /// <summary>
        /// A compile result containing a web image or a reference to a web image
        /// </summary>
        WebImage = 5
    }
}
