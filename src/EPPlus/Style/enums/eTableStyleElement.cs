/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/05/2021         EPPlus Software AB       EPPlus 5.6 
 *************************************************************************************************/
namespace OfficeOpenXml.Style
{
    public enum eTableStyleElement
    {
        /// <summary>
        /// Style that applies to a table's first column.
        /// </summary>
        FirstColumn,
        /// <summary>
        /// Style that applies to a table's first column stripes.
        /// </summary>
        FirstColumnStripe,
        /// <summary>
        /// Style that applies to a table's first header row cell.
        /// </summary>
        FirstHeaderCell,
        /// <summary>
        /// Style that applies to a table's first row stripes.
        /// </summary>
        FirstRowStripe,
        /// <summary>
        /// Style that applies to a table's first total row cell.
        /// </summary>
        FirstTotalCell,
        /// <summary>
        /// Style that applies to a table's header row.
        /// </summary>
        HeaderRow,
        /// <summary>
        /// Style that applies to a table's last column.
        /// </summary>
        LastColumn,
        /// <summary>
        /// Style that applies to a table's last header row cell.
        /// </summary>
        LastHeaderCell,
        /// <summary>
        /// Style that applies to a table's last total row cell.
        /// </summary>
        LastTotalCell,
        /// <summary>
        /// Style that applies to a table's second column stripes.
        /// </summary>
        SecondColumnStripe,
        /// <summary>
        /// Style that applies to a table's second row stripes.
        /// </summary>
        SecondRowStripe,
        /// <summary>
        /// Style that applies to a table's total row.
        /// </summary>
        TotalRow,
        /// <summary>
        /// Style that applies to a table's entire content.
        /// </summary>
        WholeTable
    }
}
