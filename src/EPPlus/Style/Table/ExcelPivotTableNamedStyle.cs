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
    public class ExcelPivotTableNamedStyle : ExcelTableNamedStyleBase
    {
        internal ExcelPivotTableNamedStyle(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles) : base(nameSpaceManager, topNode, styles)
        {
            //PageFieldLabels = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.PageFieldLabels);
            //PageFieldValues = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.PageFieldValues);
            //FirstSubtotalColumn = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.FirstSubtotalColumn);
            //SecondSubtotalColumn = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.SecondSubtotalColumn);
            //ThirdSubtotalColumn = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.ThirdSubtotalColumn);
            //BlankRow = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.BlankRow);
            //FirstSubtotalRow = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.FirstSubtotalRow);
            //SecondSubtotalRow = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.SecondSubtotalRow);
            //ThirdSubtotalRow = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.ThirdSubtotalRow);
            //FirstColumnSubheading = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.FirstColumnSubheading);
            //SecondColumnSubheading = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.SecondColumnSubheading);
            //ThirdColumnSubheading = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.ThirdColumnSubheading);
            //FirstRowSubheading = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.FirstRowSubheading);
            //SecondRowSubheading = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.SecondRowSubheading);
            //ThirdRowSubheading = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.ThirdRowSubheading);
            //GrandTotalColumn = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.LastColumn);
            //GrandTotalRow = new ExcelTableStyleElement(nameSpaceManager, topNode, styles, eTableStyleElement.TotalRow);
        }
        public ExcelTableStyleElement PageFieldLabels
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.PageFieldLabels, false);
            }
        }
        public ExcelTableStyleElement PageFieldValues
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.PageFieldValues, false);
            }
        }
        public ExcelTableStyleElement FirstSubtotalColumn
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.FirstSubtotalColumn, false);
            }
        }
        public ExcelTableStyleElement SecondSubtotalColumn
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.SecondSubtotalColumn, false);
            }
        }
        public ExcelTableStyleElement ThirdSubtotalColumn
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.ThirdSubtotalColumn, false);
            }
        }
        public ExcelTableStyleElement BlankRow
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.BlankRow, false);
            }
        }
        public ExcelTableStyleElement FirstSubtotalRow
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.FirstSubtotalRow, false);
            }
        }
        public ExcelTableStyleElement SecondSubtotalRow
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.SecondSubtotalRow, false);
            }
        }
        public ExcelTableStyleElement ThirdSubtotalRow
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.ThirdSubtotalRow, false);
            }
        }
        public ExcelTableStyleElement FirstColumnSubheading
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.FirstColumnSubheading, false);
            }
        }
        public ExcelTableStyleElement SecondColumnSubheading
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.SecondColumnSubheading, false);
            }
        }
        public ExcelTableStyleElement ThirdColumnSubheading
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.ThirdColumnSubheading, false);
            }
        }
        public ExcelTableStyleElement FirstRowSubheading
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.FirstRowSubheading, false);
            }
        }
        public ExcelTableStyleElement SecondRowSubheading
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.SecondRowSubheading, false);
            }
        }
        public ExcelTableStyleElement ThirdRowSubheading
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.ThirdRowSubheading, false);
            }
        }
        //public ExcelTableStyleElement GrandTotalColumn
        //{
        //    get
        //    {
        //        return GetTableStyleElement(eTableStyleElement.GrandTotalColumn, false);
        //    }
        //}
        //public ExcelTableStyleElement GrandTotalRow
        //{
        //    get
        //    {
        //        return GetTableStyleElement(eTableStyleElement.GrandTotalRow, false);
        //    }
        //}
    }
}
