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
    /// <summary>
    /// Custom style element for a pivot table
    /// </summary>
    public enum eTableStyleElement
    {
        /// <summary>
        /// Style that applies to a pivot table's blank rows.
        /// </summary>
        BlankRow,
        /// <summary>
        /// Style that applies to a pivot table's first column.
        /// </summary>
        FirstColumn,
        /// <summary>
        /// Style that applies to a pivot table's first column stripes.
        /// </summary>
        FirstColumnStripe,
        /// <summary>
        /// Style that applies to a pivot table's first column subheading.
        /// </summary>
        FirstColumnSubheading,
        /// <summary>
        /// Style that applies to a pivot table's first header row cell.
        /// </summary>
        FirstHeaderCell,
        /// <summary>
        /// Style that applies to a pivot table's first row stripes.
        /// </summary>
        FirstRowStripe,
        /// <summary>
        /// Style that applies to a pivot table's first row subheading.
        /// </summary>
        FirstRowSubheading,
        /// <summary>
        /// Style that applies to a pivot table's first subtotal column.
        /// </summary>
        FirstSubtotalColumn,
        /// <summary>
        /// Style that applies to a pivot table's first subtotal row.
        /// </summary>
        FirstSubtotalRow,
        /// <summary>
        /// Style that applies to a pivot table's header row.
        /// </summary>
        HeaderRow,
        /// <summary>
        /// Style that applies to a pivot table's last column.
        /// </summary>
        LastColumn,
        /// <summary>
        /// Style that applies to a pivot table's page field labels.
        /// </summary>
        PageFieldLabels,
        /// <summary>
        /// Style that applies to a pivot table's page field values.
        /// </summary>
        PageFieldValues,
        /// <summary>
        /// Style that applies to a pivot table's second column stripes.
        /// </summary>
        SecondColumnStripe,
        /// <summary>
        /// Style that applies to a pivot table's second column subheading.
        /// </summary>
        SecondColumnSubheading,
        /// <summary>
        /// Style that applies to a pivot table's second row stripes.
        /// </summary>
        SecondRowStripe,
        /// <summary>
        /// Style that applies to a pivot table's second row subheading.
        /// </summary>
        SecondRowSubheading,
        /// <summary>
        /// Style that applies to a pivot table's second subtotal column.
        /// </summary>
        SecondSubtotalColumn,
        /// <summary>
        /// Style that applies to a pivot table's second subtotal row.
        /// </summary>
        SecondSubtotalRow,
        /// <summary>
        /// Style that applies to a pivot table's third column subheading.
        /// </summary>
        ThirdColumnSubheading,
        /// <summary>
        /// Style that applies to a pivot table's third row subheading.
        /// </summary>
        ThirdRowSubheading,
        /// <summary>
        /// Style that applies to a pivot table's third subtotal column.
        /// </summary>
        ThirdSubtotalColumn,
        /// <summary>
        /// Style that applies to a pivot table's third subtotal row.
        /// </summary>
        ThirdSubtotalRow,
        /// <summary>
        /// Style that applies to a pivot table's total row.
        /// </summary>
        TotalRow,
        /// <summary>
        /// Style that applies to a pivot table's entire content.
        /// </summary>
        WholeTable,
        /// <summary>
        /// Style that applies to a table's last header row cell.
        /// </summary>
        LastHeaderCell,
        /// <summary>
        /// Style that applies to a table's first total row cell.
        /// </summary>
        FirstTotalCell,
        /// <summary>
        /// Style that applies to a table's last total row cell.
        /// </summary>
        LastTotalCell
    }
}
