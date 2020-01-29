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
        /// Type of system where the VBA project was created.
        /// </summary>
        public enum eSyskind
        {
            /// <summary>
            /// Windows 16-bit
            /// </summary>
            Win16 = 0,
            /// <summary>
            /// Windows 32-bit
            /// </summary>
            Win32 = 1,
            /// <summary>
            /// Mac
            /// </summary>
            Macintosh = 2,
            /// <summary>
            /// Windows 64-bit
            /// </summary>
            Win64 = 3
        }
}