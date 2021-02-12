/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/28/2020         EPPlus Software AB       Pivot Table Styling - EPPlus 5.6
 *************************************************************************************************/
using OfficeOpenXml.Core;
using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotAreaReference : XmlHelper
    {
        [Flags]
        internal enum ePivotSubTotalFunction
        {
            AvgSubtotal = 0x1,
            CountASubtotal = 0x2,
            CountSubtotal = 0x4,
            MaxSubtotal = 0x8,
            MinSubtotal = 0x10,
            ProductSubtotal = 0x20,
            StdDevPSubtotal = 0x40,
            StdDevSubtotal = 0x80,
            SumSubtotal = 0x100,
            VarPSubtotal = 0x200,
            VarSubtotal = 0x400
        }
        ExcelPivotTable _pt;
        internal ExcelPivotAreaReference(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt, int fieldIndex=-1) : base(nsm, topNode)
        {
            _pt = pt;
            if (fieldIndex >= 0)
            {
                FieldIndex = fieldIndex;
                var cache=Field.Cache;
                var items = cache.SharedItems.Count > 0 ? cache.SharedItems : cache.GroupItems;
                foreach (XmlNode n in topNode.ChildNodes)
                {
                    if (n.LocalName == "x")
                    {
                        var ix = int.Parse(n.Attributes["x"].Value);
                        if (ix < items.Count)
                        {
                            Values.Add(items[ix]);
                        }
                    }
                }
            }
            else if(fieldIndex==-2)
            {
                SetXmlNodeLong("@field", 4294967294);
            }
        }
        internal int FieldIndex
        { 
            get
            {
                return GetXmlNodeInt("@field");
            }
            set
            {
                SetXmlNodeInt("@field", value);
            }
        }
        /// <summary>
        /// The pivot table field referenced
        /// </summary>
        public ExcelPivotTableField Field 
        { 
            get
            {
                if(FieldIndex >= 0)
                {
                    return _pt.Fields[FieldIndex];
                }
                return null;
            }
        }
        public bool Selected { get; set; } = true;
        internal bool Relative 
        { 
            get
            {
                return GetXmlNodeBool("@relative");
            }
            set
            {
                SetXmlNodeBool("@relative", value);
            }
        }
        internal bool ByPosition 
        {
            get
            {
                return GetXmlNodeBool("@byPosition");
            }
            set
            {
                SetXmlNodeBool("@byPosition", value);
            }
        }
        public List<object> Values { get; } = new List<object>();
        public bool DefaultSubtotal 
        { 
            get
            {
                return GetXmlNodeBool("@defaultSubtotal");
            }
            set
            {
                SetXmlNodeBool("@defaultSubtotal", value);
            }
        }
        public bool AvgSubtotal
        {
            get
            {
                return GetXmlNodeBool("@avgSubtotal");
            }
            set
            {
                SetXmlNodeBool("@avgSubtotal", value);
            }
        }
        public bool CountSubtotal
        {
            get
            {
                return GetXmlNodeBool("@countSubtotal");
            }
            set
            {
                SetXmlNodeBool("@countSubtotal", value);
            }
        }
        public bool CountASubtotal
        {
            get
            {
                return GetXmlNodeBool("@countASubtotal");
            }
            set
            {
                SetXmlNodeBool("@countASubtotal", value);
            }
        }
        public bool MaxSubtotal
        {
            get
            {
                return GetXmlNodeBool("@maxSubtotal");
            }
            set
            {
                SetXmlNodeBool("@maxSubtotal", value);
            }
        }
        public bool MinSubtotal
        {
            get
            {
                return GetXmlNodeBool("@minSubtotal");
            }
            set
            {
                SetXmlNodeBool("@minSubtotal", value);
            }
        }
        public bool ProductSubtotal
        {
            get
            {
                return GetXmlNodeBool("@productSubtotal");
            }
            set
            {
                SetXmlNodeBool("@productSubtotal", value);
            }
        }
        public bool StdDevPSubtotal
        {
            get
            {
                return GetXmlNodeBool("@StdDevPSubtotal");
            }
            set
            {
                SetXmlNodeBool("@StdDevPSubtotal", value);
            }
        }
        public bool StdDevSubtotal
        {
            get
            {
                return GetXmlNodeBool("@StdDevSubtotal");
            }
            set
            {
                SetXmlNodeBool("@StdDevSubtotal", value);
            }
        }
        public bool SumSubtotal
        {
            get
            {
                return GetXmlNodeBool("@sumSubtotal");
            }
            set
            {
                SetXmlNodeBool("@sumSubtotal", value);
            }
        }
        public bool VarPSubtotal
        {
            get
            {
                return GetXmlNodeBool("@varPSubtotal");
            }
            set
            {
                SetXmlNodeBool("@varPSubtotal", value);
            }
        }
        public bool VarSubtotal
        {
            get
            {
                return GetXmlNodeBool("@varSubtotal");
            }
            set
            {
                SetXmlNodeBool("@varSubtotal", value);
            }
        }
        internal void SetFunction(DataFieldFunctions function)
        {
            switch(function)
            {
                case DataFieldFunctions.Average:
                    AvgSubtotal = true;
                    break;
                case DataFieldFunctions.Count:
                    CountSubtotal = true;
                    break;
                case DataFieldFunctions.CountNums:
                    CountASubtotal = true;
                    break;
                case DataFieldFunctions.Max:
                    MaxSubtotal = true;
                    break;
                case DataFieldFunctions.Min:
                    MinSubtotal = true;
                    break;
                case DataFieldFunctions.Product:
                    ProductSubtotal = true;
                    break;
                case DataFieldFunctions.StdDevP:
                    StdDevPSubtotal = true;
                    break;
                case DataFieldFunctions.StdDev:
                    StdDevSubtotal = true;
                    break;
                case DataFieldFunctions.Sum:
                    SumSubtotal = true;
                    break;
                case DataFieldFunctions.VarP:
                    VarPSubtotal = true;
                    break;
                case DataFieldFunctions.Var:
                    VarSubtotal = true;
                    break;
                default:
                    DefaultSubtotal = true;
                    break;
            }
        }
    }
}