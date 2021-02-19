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
using System.Collections;

namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotAreaDataFieldReference : ExcelPivotAreaReferenceBase, IEnumerable<ExcelPivotTableDataField>
    {
        List<ExcelPivotTableDataField> _dataFields = new List<ExcelPivotTableDataField>();
        internal ExcelPivotAreaDataFieldReference(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt, int fieldIndex = -1) : base(nsm, topNode, pt)
        {
            if(TopNode.LocalName=="reference")
            {
                foreach (XmlNode c in TopNode.ChildNodes)
                {
                    if (c.LocalName == "x")
                    {
                        var ix = int.Parse(c.Attributes["v"].Value);
                        if (ix < pt.DataFields.Count)
                        {
                            _dataFields.Add(pt.DataFields[ix]);
                        }
                    }
                }
            }
        }
        public ExcelPivotTableDataField this[int index]
        {
            get
            {
                return _dataFields[index];
            }
        }
        public int Count 
        { 
            get
            {
                return _dataFields.Count;
            }
        }
        internal void AddInternal(ExcelPivotTableDataField item)
        {
            _dataFields.Add(item);
        }
        public void Add(int index)
        {
            if (index >= 0 && index < _pt.DataFields.Count)
            {
                _dataFields.Add(_pt.DataFields[index]);
            }
            else
            {
                throw new IndexOutOfRangeException("Index is out of range for referenced data field.");
            }
        }
        public void Add(ExcelPivotTableDataField field)
        {
            if (field == null)
            {
                throw new ArgumentNullException("The pivot table field must not be null.");
            }
            if (field.Field._pivotTable != _pt)
            {
                throw new ArgumentException("The pivot table field is from another pivot table.");
            }
            _dataFields.Add(field);
        }

        internal override void UpdateXml()
        {
            if(_dataFields.Count==0)
            {
                if(TopNode.LocalName == "reference")
                {
                    TopNode.ParentNode.RemoveChild(TopNode);
                }
                return;
            }
            else
            {
                if (TopNode.LocalName == "pivotArea")
                {
                    var n = CreateNode("d:references");
                    var rn = (XmlElement)CreateNode(n, "d:reference", true);
                    rn.SetAttribute("field", "4294967294");
                    TopNode = rn;
                }
            }

            foreach (ExcelPivotTableDataField r in _dataFields)
            {
                if (r.Field.IsDataField)
                {
                    var ix = _pt.DataFields._list.IndexOf(r);
                    if (ix >= 0)
                    {
                        var n = (XmlElement)CreateNode("d:x", false, true);
                        n.SetAttribute("v", ix.ToString(CultureInfo.InvariantCulture));
                    }
                }
            }            
        }

        public IEnumerator<ExcelPivotTableDataField> GetEnumerator()
        {
            return ((IEnumerable<ExcelPivotTableDataField>)_dataFields).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable)_dataFields).GetEnumerator();
        }
    }
}