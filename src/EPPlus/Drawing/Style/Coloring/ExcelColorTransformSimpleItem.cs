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
    /// A tranformation operation for a color
    /// </summary>
    internal class ExcelColorTransformSimpleItem : IColorTransformItem, ISource
    {
        /// <summary>
        /// Type of tranformation
        /// </summary>
        public eColorTransformType Type { get; set; }

        /// <summary>
        /// The datatype of the value
        /// </summary>
        public eColorTransformDataType DataType { get; set; }
        /// <summary>
        /// The value
        /// </summary>
        public double Value { get; set; }

        bool ISource._fromStyleTemplate { get; set; } = false;
    }
}