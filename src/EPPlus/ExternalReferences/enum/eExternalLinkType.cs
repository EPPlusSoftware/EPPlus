/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/28/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
namespace OfficeOpenXml.ExternalReferences
{
    /// <summary>
    /// The type of external link
    /// </summary>
    public enum eExternalLinkType
    {
        /// <summary>
        /// The external link is of type <see cref="ExcelExternalWorkbook" />
        /// </summary>
        ExternalWorkbook,
        /// <summary>
        /// The external link is of type <see cref="ExcelExternalDdeLink" />
        /// </summary>
        DdeLink,
        /// <summary>
        /// The external link is of type <see cref="ExcelExternalOleLink" />
        /// </summary>
        OleLink
    }
}