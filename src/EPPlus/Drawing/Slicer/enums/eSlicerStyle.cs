/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Slicer;

namespace OfficeOpenXml
{
    /// <summary>
    /// Buildin slicer styles
    /// </summary>
    public enum eSlicerStyle
    {
        /// <summary>
        /// No slicer style specified
        /// </summary>
        None,
        /// <summary>
        /// A custom style set by the <see cref="ExcelSlicer.Style"/> property
        /// </summary>
        Custom,
        /// <summary>
        /// Light 1 style
        /// </summary>
        Light1,
        /// <summary>
        /// Light 2 style
        /// </summary>
        Light2,
        /// <summary>
        /// Light 3 style
        /// </summary>
        Light3,
        /// <summary>
        /// Light 4 style
        /// </summary>
        Light4,
        /// <summary>
        /// Light 5 style
        /// </summary>
        Light5,
        /// <summary>
        /// Light 6 style
        /// </summary>
        Light6,
        /// <summary>
        /// Other 1 style
        /// </summary>
        Other1,
        /// <summary>
        /// Other 2 style
        /// </summary>
        Other2,
        /// <summary>
        /// Dark 1 style
        /// </summary>
        Dark1,
        /// <summary>
        /// Dark 2 style
        /// </summary>
        Dark2,
        /// <summary>
        /// Dark 3 style
        /// </summary>
        Dark3,
        /// <summary>
        /// Dark 4 style
        /// </summary>
        Dark4,
        /// <summary>
        /// Dark 5 style
        /// </summary>
        Dark5,
        /// <summary>
        /// Dark 6 style
        /// </summary>
        Dark6
    }
}