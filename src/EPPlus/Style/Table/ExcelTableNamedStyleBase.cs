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
using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Style.Table
{
    public abstract class ExcelTableNamedStyleBase : XmlHelper
    {
        ExcelStyles _styles;
        internal Dictionary<eTableStyleElement, ExcelTableStyleElement> _dic = new Dictionary<eTableStyleElement, ExcelTableStyleElement>();
        internal ExcelTableNamedStyleBase(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles) : base(nameSpaceManager, topNode)
        {
            _styles = styles;
            As = new ExcelTableNamedStyleAsType(this);
        }
        protected ExcelTableStyleElement GetTableStyleElement(eTableStyleElement element, bool createBanded)
        {
            if (_dic.ContainsKey(element))
            {
                return _dic[element];
            }
            ExcelTableStyleElement item;
            if (createBanded)
            {
                item = new ExcelBandedTableStyleElement(NameSpaceManager, TopNode, _styles, element);
            }
            else
            {
                item = new ExcelTableStyleElement(NameSpaceManager, TopNode, _styles, element);
            }
            _dic.Add(element, item);
            return item;
        }
        public abstract eTableNamedStyleAppliesTo AppliesTo
        {
            get;
        }
        /// <summary>
        /// The name of the table named style
        /// </summary>
        public string Name 
        { 
            get
            {
                return GetXmlNodeString("@name");
            }
            set
            {
                if(_styles.TableStyles.ExistsKey(value))
                {
                    throw new InvalidOperationException("Name already exists in the collection");
                }
                SetXmlNodeString("@name", value);
            }
        }
        /// <summary>
        /// Applies to the entire content of a table or pivot table
        /// </summary>
        public ExcelTableStyleElement WholeTable 
        { 
            get
            {
                return GetTableStyleElement(eTableStyleElement.WholeTable, false);
            }
        }
        /// <summary>
        /// Applies to the first column stripe of a table or pivot table
        /// </summary>
        public ExcelBandedTableStyleElement FirstColumnStripe
        { 
            get
            {
                return (ExcelBandedTableStyleElement)GetTableStyleElement(eTableStyleElement.FirstColumnStripe, true);
            }
        }
        /// <summary>
        /// Applies to the second column stripe of a table or pivot table
        /// </summary>
        public ExcelBandedTableStyleElement SecondColumnStripe
        {
            get
            {
                return (ExcelBandedTableStyleElement)GetTableStyleElement(eTableStyleElement.SecondColumnStripe, true);
            }
        }
        /// <summary>
        /// Applies to the first row stripe of a table or pivot table
        /// </summary>
        public ExcelBandedTableStyleElement FirstRowStripe
        {
            get
            {
                return (ExcelBandedTableStyleElement)GetTableStyleElement(eTableStyleElement.FirstRowStripe, true);
            }
        }
        /// <summary>
        /// Applies to the second row stripe of a table or pivot table
        /// </summary>
        public ExcelBandedTableStyleElement SecondRowStripe
        {
            get
            {
                return (ExcelBandedTableStyleElement)GetTableStyleElement(eTableStyleElement.SecondRowStripe, true);
            }
        }
        /// <summary>
        /// Applies to the last column of a table or pivot table
        /// </summary>
        public ExcelTableStyleElement LastColumn
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.LastColumn, false);
            }
        }
        /// <summary>
        /// Applies to the first column of a table or pivot table
        /// </summary>
        public ExcelTableStyleElement FirstColumn
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.FirstColumn, false);
            }
        }
        /// <summary>
        /// Applies to the header row of a table or pivot table
        /// </summary>
        public ExcelTableStyleElement HeaderRow
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.HeaderRow, false);
            }
        }
        /// <summary>
        /// Applies to the total row of a table or pivot table
        /// </summary>
        public ExcelTableStyleElement TotalRow
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.TotalRow, false);
            }
        }
        /// <summary>
        /// Applies to the first header cell of a table or pivot table
        /// </summary>
        public ExcelTableStyleElement FirstHeaderCell
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.FirstHeaderCell, false);
            }
        }
        /// <summary>
        /// Provides access to type conversion for all table named styles.
        /// </summary>
        public ExcelTableNamedStyleAsType As
        {
            get;
        }
    }
}

