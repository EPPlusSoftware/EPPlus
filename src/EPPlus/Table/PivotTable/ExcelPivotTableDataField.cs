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
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A pivot table data field
    /// </summary>
    public class ExcelPivotTableDataField : XmlHelper
    {
        internal ExcelPivotTableDataField(XmlNamespaceManager ns, XmlNode topNode,ExcelPivotTableField field) :
            base(ns, topNode)
        {
            if (topNode.Attributes.Count == 0)
            {
                Index = field.Index;
                BaseField = 0;
                BaseItem = 0;
            }
            
            Field = field;
        }
        /// <summary>
        /// The field
        /// </summary>
        public ExcelPivotTableField Field
        {
            get;
            internal set;
        }
        /// <summary>
        /// The index of the datafield
        /// </summary>
        public int Index 
        { 
            get
            {
                return GetXmlNodeInt("@fld");
            }
            internal set
            {
                SetXmlNodeString("@fld",value.ToString());
            }
        }
        /// <summary>
        /// The name of the datafield
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString("@name");
            }
            set
            {
                if (Field._pivotTable.DataFields.ExistsDfName(value, this))
                {
                    throw (new InvalidOperationException("Duplicate datafield name"));
                }
                SetXmlNodeString("@name", value);
            }
        }
        /// <summary>
        /// Field index. Reference to the field collection
        /// </summary>
        public int BaseField
        {
            get
            {
                return GetXmlNodeInt("@baseField");
            }
            set
            {
                SetXmlNodeString("@baseField", value.ToString());
            }
        }
        /// <summary>
        /// The index to the base item when the ShowDataAs calculation is in use
        /// </summary>
        public int BaseItem
        {
            get
            {
                return GetXmlNodeInt("@baseItem");
            }
            set
            {
                SetXmlNodeString("@baseItem", value.ToString());
            }
        }
        /// <summary>
        /// Number format id. 
        /// </summary>
        internal int NumFmtId
        {
            get
            {
                return GetXmlNodeInt("@numFmtId");
            }
            set
            {
                SetXmlNodeString("@numFmtId", value.ToString());
            }
        }
        /// <summary>
        /// The number format for the data field
        /// </summary>
        public string Format
        {
            get
            {
                foreach (var nf in Field._pivotTable.WorkSheet.Workbook.Styles.NumberFormats)
                {
                    if (nf.NumFmtId == NumFmtId)
                    {
                        return nf.Format;
                    }
                }
                return Field._pivotTable.WorkSheet.Workbook.Styles.NumberFormats[0].Format;
            }
            set
            {
                var styles = Field._pivotTable.WorkSheet.Workbook.Styles;

                ExcelNumberFormatXml nf = null;
                if (!styles.NumberFormats.FindById(value, ref nf))
                {
                    nf = new ExcelNumberFormatXml(NameSpaceManager) { Format = value, NumFmtId = styles.NumberFormats.NextId++ };
                    styles.NumberFormats.Add(value, nf);
                }
                NumFmtId = nf.NumFmtId;
            }
        }
        /// <summary>
        /// Type of aggregate function
        /// </summary>
        public DataFieldFunctions Function
        {
            get
            {
                string s=GetXmlNodeString("@subtotal");
                if(s=="")
                {
                    return DataFieldFunctions.None;
                }
                else
                {
                    return (DataFieldFunctions)Enum.Parse(typeof(DataFieldFunctions), s, true);
                }
            }
            set
            {
                string v;
                switch(value)
                {
                    case DataFieldFunctions.None:
                        DeleteNode("@subtotal");
                        return;
                    case DataFieldFunctions.CountNums:
                        v="countNums";
                        break;
                    case DataFieldFunctions.StdDev:
                        v="stdDev";
                        break;
                    case DataFieldFunctions.StdDevP:
                        v="stdDevP";
                        break;
                    default:
                        v=value.ToString().ToLower(CultureInfo.InvariantCulture);
                        break;
                }                
                SetXmlNodeString("@subtotal", v);
            }
        }
        ExcelPivotTableDataFieldShowDataAs _showDataAs = null;
        public ExcelPivotTableDataFieldShowDataAs ShowDataAs
        {
            get
            {
                if (_showDataAs == null)
                {
                    _showDataAs = new ExcelPivotTableDataFieldShowDataAs(this);
                }
                return _showDataAs;
            }
        }
        internal eShowDataAs ShowDataAsInternal
        {
            get
            {
                string s = GetXmlNodeString("@showDataAs");
                if (s == "")
                {
                    s = GetXmlNodeString("d:extLst/d:ext[@uri='{E15A36E0-9728-4e99-A89B-3F7291B0FE68}']/x14:dataField/@pivotShowAs");
                    if (s == "")
                    {
                        return eShowDataAs.Normal;
                    }
                }
                return s.ToShowDataAs();
            }
            set
            {
                if(value==eShowDataAs.Normal)
                {
                    DeleteNode("@showDataAs");
                }
                else
                {
                    if(IsShowDataAsExtLst(value))
                    {
                        DeleteNode("@showDataAs");
                        var extNode = GetOrCreateExtLstSubNode("{E15A36E0-9728-4e99-A89B-3F7291B0FE68}", "x14");
                        var extNodeHelper = XmlHelperFactory.Create(NameSpaceManager, extNode);

                        extNodeHelper.SetXmlNodeString("x14:dataField/@pivotShowAs", value.FromShowDataAs());
                    }
                    else
                    {
                        DeleteNode("d:extLst/d:ext[@url='{E15A36E0-9728-4e99-A89B-3F7291B0FE68}']");
                        SetXmlNodeString("@showDataAs", value.FromShowDataAs());
                    }
                }
            }
        }

        private bool IsShowDataAsExtLst(eShowDataAs value)
        {
            return
               value == eShowDataAs.PercentOfParent ||
               value == eShowDataAs.PercentOfParentColumn ||
               value == eShowDataAs.PercentOfParentRow ||
               value == eShowDataAs.RankAscending ||
               value == eShowDataAs.RankDescending ||
               value == eShowDataAs.PercentOfRunningTotal;
        }
    }
}
