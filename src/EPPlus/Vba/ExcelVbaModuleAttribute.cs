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
namespace OfficeOpenXml.VBA
{
    /// <summary>
    /// A VBA modual attribute
    /// </summary>
    public class ExcelVbaModuleAttribute
    {
        internal ExcelVbaModuleAttribute()
        {

        }
        /// <summary>
        /// The name of the attribute
        /// </summary>
        public string Name { get; internal set; }
        /// <summary>
        /// The datatype. Determine if the attribute uses double quotes around the value.
        /// </summary>
        public eAttributeDataType DataType { get; internal set; }
        /// <summary>
        /// The value of the attribute without any double quotes.
        /// </summary>
        public string Value { get; set; }
        /// <summary>
        /// Converts the object to a string
        /// </summary>
        /// <returns>The name of the VBA module attribute</returns>
        public override string ToString()
        {
            return Name;
        }
    }
}
