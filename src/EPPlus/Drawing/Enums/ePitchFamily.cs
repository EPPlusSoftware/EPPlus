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
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Specifies the font pitch
    /// </summary>
    public enum ePitchFamily
    {
        /// <summary>
        /// Default pitch + unknown font family
        /// </summary>
        Default = 0x00, 
        /// <summary>
        /// Fixed pitch + unknown font family
        /// </summary>
        Fixed = 0x01, 
        /// <summary>
        /// Variable pitch + unknown font family
        /// </summary>
        Variable = 0x02,
        /// <summary>
        /// Default pitch + Roman font family
        /// </summary>
        DefaultRoman = 0x10,    
        /// <summary>
        /// Fixed pitch + Roman font family
        /// </summary>
        FixedRoman = 0x11,      
        /// <summary>
        /// Variable pitch + Roman font family
        /// </summary>
        VariableRoman = 0x12,   
        /// <summary>
        /// Default pitch + Swiss font family
        /// </summary>
        DefaultSwiss = 0x20,    
        /// <summary>
        /// Fixed pitch + Swiss font family
        /// </summary>
        FixedSwiss = 0x21,      
        /// <summary>
        /// Variable pitch + Swiss font family
        /// </summary>
        VariableSwiss = 0x22 
    }
}