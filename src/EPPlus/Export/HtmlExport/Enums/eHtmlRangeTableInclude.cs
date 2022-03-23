/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/11/2021         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
namespace OfficeOpenXml
{
    /// <summary>
    /// How to handle a range when it is a table.
    /// </summary>
    public enum eHtmlRangeTableInclude
    {
        /// <summary>
        /// Do not set the table style css classes on the html table or create the table style css.
        /// </summary>
        Exclude,
        /// <summary>
        /// Set the css table style classes on the table, but do not include the table classes in the css.
        /// </summary>
        ClassNamesOnly,
        /// <summary>
        /// Include the css table style for the table and set the corresponding classes on the html table.
        /// </summary>
        Include
    }
}