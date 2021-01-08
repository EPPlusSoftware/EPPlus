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
        }
        protected ExcelTableStyleElement GetTableStyleElement(eTableStyleElement element, bool createBanded)
        {
            if(_dic.ContainsKey(element))
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
        public ExcelTableStyleElement WholeTable 
        { 
            get
            {
                return GetTableStyleElement(eTableStyleElement.WholeTable, false);
            }
        }


        public ExcelBandedTableStyleElement FirstColumnStripe
        { 
            get
            {
                return (ExcelBandedTableStyleElement)GetTableStyleElement(eTableStyleElement.FirstColumnStripe, true);
            }
        }
        public ExcelBandedTableStyleElement SecondColumnStripe
        {
            get
            {
                return (ExcelBandedTableStyleElement)GetTableStyleElement(eTableStyleElement.SecondColumnStripe, true);
            }
        }
        public ExcelBandedTableStyleElement FirstRowStripe
        {
            get
            {
                return (ExcelBandedTableStyleElement)GetTableStyleElement(eTableStyleElement.FirstRowStripe, true);
            }
        }
        public ExcelBandedTableStyleElement SecondRowStripe
        {
            get
            {
                return (ExcelBandedTableStyleElement)GetTableStyleElement(eTableStyleElement.SecondRowStripe, true);
            }
        }
        public ExcelTableStyleElement LastColumn
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.LastColumn, false);
            }
        }
        public ExcelTableStyleElement FirstColumn
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.FirstColumn, false);
            }
        }
        public ExcelTableStyleElement HeaderRow
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.HeaderRow, false);
            }
        }
        public ExcelTableStyleElement TotalRow
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.TotalRow, false);
            }
        }
        public ExcelTableStyleElement FirstHeaderCell
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.FirstHeaderCell, false);
            }
        }
    }
}
