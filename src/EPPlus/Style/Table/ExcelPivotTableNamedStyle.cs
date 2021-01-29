/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/08/2021         EPPlus Software AB       Table Styling - EPPlus 5.6
 *************************************************************************************************/
using System.Xml;

namespace OfficeOpenXml.Style.Table
{
    /// <summary>
    /// A custom named table style that applies to pivot tables only
    /// </summary>
    public class ExcelPivotTableNamedStyle : ExcelTableNamedStyleBase
    {
        internal ExcelPivotTableNamedStyle(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles) : base(nameSpaceManager, topNode, styles)
        {
        }
        /// <summary>
        /// If the style applies to tables, pivot table or both
        /// </summary>
        public override eTableNamedStyleAppliesTo AppliesTo
        {
            get
            {
                return eTableNamedStyleAppliesTo.PivotTables;
            }
        }

        /// <summary>
        /// Applies to the page field labels of a pivot table
        /// </summary>
        public ExcelTableStyleElement PageFieldLabels
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.PageFieldLabels);
            }
        }
        /// <summary>
        /// Applies to the page field values of a pivot table
        /// </summary>
        public ExcelTableStyleElement PageFieldValues
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.PageFieldValues);
            }
        }
        /// <summary>
        /// Applies to the first subtotal column of a pivot table
        /// </summary>
        public ExcelTableStyleElement FirstSubtotalColumn
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.FirstSubtotalColumn);
            }
        }
        /// <summary>
        /// Applies to the second subtotal column of a pivot table
        /// </summary>
        public ExcelTableStyleElement SecondSubtotalColumn
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.SecondSubtotalColumn);
            }
        }
        /// <summary>
        /// Applies to the third subtotal column of a pivot table
        /// </summary>
        public ExcelTableStyleElement ThirdSubtotalColumn
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.ThirdSubtotalColumn);
            }
        }
        /// <summary>
        /// Applies to blank rows of a pivot table
        /// </summary>
        public ExcelTableStyleElement BlankRow
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.BlankRow);
            }
        }
        /// <summary>
        /// Applies to the first subtotal row of a pivot table
        /// </summary>
        public ExcelTableStyleElement FirstSubtotalRow
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.FirstSubtotalRow);
            }
        }
        /// <summary>
        /// Applies to the second subtotal row of a pivot table
        /// </summary>
        public ExcelTableStyleElement SecondSubtotalRow
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.SecondSubtotalRow);
            }
        }
        /// <summary>
        /// Applies to the third subtotal row of a pivot table
        /// </summary>
        public ExcelTableStyleElement ThirdSubtotalRow
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.ThirdSubtotalRow);
            }
        }
        /// <summary>
        /// Applies to the first column subheading of a pivot table
        /// </summary>
        public ExcelTableStyleElement FirstColumnSubheading
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.FirstColumnSubheading);
            }
        }
        /// <summary>
        /// Applies to the second column subheading of a pivot table
        /// </summary>
        public ExcelTableStyleElement SecondColumnSubheading
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.SecondColumnSubheading);
            }
        }
        /// <summary>
        /// Applies to the third column subheading of a pivot table
        /// </summary>
        public ExcelTableStyleElement ThirdColumnSubheading
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.ThirdColumnSubheading);
            }
        }
        /// <summary>
        /// Applies to the first row subheading of a pivot table
        /// </summary>
        public ExcelTableStyleElement FirstRowSubheading
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.FirstRowSubheading);
            }
        }
        /// <summary>
        /// Applies to the second row subheading of a pivot table
        /// </summary>
        public ExcelTableStyleElement SecondRowSubheading
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.SecondRowSubheading);
            }
        }
        /// <summary>
        /// Applies to the third row subheading of a pivot table
        /// </summary>
        public ExcelTableStyleElement ThirdRowSubheading
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.ThirdRowSubheading);
            }
        }
    }
}
