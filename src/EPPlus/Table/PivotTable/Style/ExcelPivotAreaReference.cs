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
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using System.Linq;
using OfficeOpenXml.Core;

namespace OfficeOpenXml.Table.PivotTable
{
    public struct PivotReference
    {
        public int Index { get; internal set; }
        public object Value { get; internal set; }
    }
    public class ExcelPivotAreaReference : ExcelPivotAreaReferenceBase
    {
        internal ExcelPivotAreaReference(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt, int fieldIndex = -1) : base(nsm, topNode, pt)
        {
            if (fieldIndex != -1)
            {
                FieldIndex = fieldIndex;
            }
            if (FieldIndex >= 0)
            {
                var cache = Field.Cache;
                var items = cache.SharedItems.Count > 0 ? cache.SharedItems : cache.GroupItems;
                foreach (XmlNode n in topNode.ChildNodes)
                {
                    if (n.LocalName == "x")
                    {
                        var ix = int.Parse(n.Attributes["v"].Value);
                        if (ix < items.Count)
                        {
                            CacheItems.Add(new PivotReference() { Index = ix, Value = items[ix] });
                        }
                    }
                }
            }
        }
        /// <summary>
        /// The pivot table field referenced
        /// </summary>
        public ExcelPivotTableField Field
        {
            get
            {
                if (FieldIndex >= 0)
                {
                    return _pt.Fields[FieldIndex];
                }
                return null;
            }
        }
        /// <summary>
        /// References to the pivot table cache or within the table.
        /// </summary>
        public List<PivotReference> CacheItems { get; } = new List<PivotReference>();
        public void AddItemByIndex(int index)
        {
            {
                var items = Field.Cache.SharedItems.Count == 0 ? Field.Cache.GroupItems : Field.Cache.SharedItems;
                if (items.Count > index)
                {
                    CacheItems.Add(new PivotReference() { Index = index, Value = items[index] });
                }
                else
                {
                    throw new IndexOutOfRangeException("Index is out of range in cache Items. Please make sure the pivot table cache has been refreshed.");
                }
            }
        }
        public void AddItemByValue(object value)
        {
            var items = Field.Cache.SharedItems.Count == 0 ? Field.Cache.GroupItems : Field.Cache.SharedItems;
            var index = items.GetIndexByValue(value);
            if (index >= 0)
            {
                CacheItems.Add(new PivotReference() { Index = index, Value = value });
            }
        }
        internal override void UpdateXml()
        {
            if (FieldIndex >= 0 && FieldIndex < _pt.Fields.Count)
            {
                var items = Field.Cache.SharedItems.Count == 0 ? Field.Cache.GroupItems : Field.Cache.SharedItems;
                foreach (var r in CacheItems)
                {
                    if (r.Index >= 0 && r.Index <= items.Count && r.Value.Equals(items[r.Index]))
                    {
                        var n = (XmlElement)CreateNode("d:x", false, true);
                        n.SetAttribute("v", r.Index.ToString(CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        var ix = items.GetIndexByValue(r.Value);
                        if (ix >= 0)
                        {
                            var n = (XmlElement)CreateNode("d:x", false, true);
                            n.SetAttribute("v", ix.ToString(CultureInfo.InvariantCulture));
                        }
                    }
                }
            }
        }
    }
    public class ExcelPivotAreaDataFieldReference : ExcelPivotAreaReferenceBase
    {
        internal ExcelPivotAreaDataFieldReference(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt, int fieldIndex = -1) : base(nsm, topNode, pt)
        {
            foreach (XmlNode n in topNode.ChildNodes)
            {
                if (n.LocalName == "x")
                {
                    var ix = int.Parse(n.Attributes["v"].Value);
                    if (ix < pt.DataFields.Count)
                    {
                        DataFields.Add(pt.DataFields[ix]);
                    }
                }
            }
        }
        public EPPlusReadOnlyList<ExcelPivotTableDataField> DataFields { get; } = new EPPlusReadOnlyList<ExcelPivotTableDataField>();
        public void AddReferenceByIndex(int index)
        {
            if (index >= 0 && index < _pt.DataFields.Count)
            {
                DataFields.Add(_pt.DataFields[index]);
            }
            else
            {
                throw new IndexOutOfRangeException("Index is out of range for referenced data field.");
            }
        }
        public void AddDataField(ExcelPivotTableDataField field)
        {
            if (field == null)
            {
                throw new ArgumentNullException("The pivot table field must not be null.");
            }
            if (field.Field._pivotTable != _pt)
            {
                throw new ArgumentException("The pivot table field is from another pivot table.");
            }
            DataFields.Add(field);
        }

        internal override void UpdateXml()
        {
            foreach (ExcelPivotTableDataField r in DataFields)
            {
                if (r.Field.IsDataField)
                {
                    var n = (XmlElement)CreateNode("d:x", false, true);
                    n.SetAttribute("v", r.Index.ToString(CultureInfo.InvariantCulture));
                }
            }            
        }
    }
    public abstract class ExcelPivotAreaReferenceBase : XmlHelper
    {
        internal protected ExcelPivotTable _pt;
        internal ExcelPivotAreaReferenceBase(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt) : base(nsm, topNode)
        {
            _pt = pt;
        }
        internal int FieldIndex
        { 
            get
            {
                var v=GetXmlNodeLong("@field");
                if(v > int.MaxValue)
                {
                    return -2;
                }
                else
                {
                    return (int)v;
                }
            }
            set
            {
                if(value<0)
                {
                    SetXmlNodeLong("@field", 4294967294);
                }
                else
                {
                    SetXmlNodeInt("@field", value);
                }
            }
        }
        public bool Selected 
        {
            get
            {
                return GetXmlNodeBool("@selected", true);
            }
            set
            {
                SetXmlNodeBool("@selected", value);
            }
        }
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
        internal abstract void UpdateXml();
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