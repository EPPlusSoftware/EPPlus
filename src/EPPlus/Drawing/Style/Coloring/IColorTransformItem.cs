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
namespace OfficeOpenXml.Drawing.Style.Coloring
{
    /// <summary>
    /// Color transformation item
    /// </summary>
    public interface IColorTransformItem
    {
        /// <summary>
        /// Type of tranformation
        /// </summary>
        eColorTransformType Type { get; }
        /// <summary>
        /// Datetype of the value property
        /// </summary>
        eColorTransformDataType DataType { get; }
        /// <summary>
        /// The value
        /// </summary>
        double Value { get; set; }
    }
}