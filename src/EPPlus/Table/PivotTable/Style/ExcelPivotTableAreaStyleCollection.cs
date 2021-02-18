using OfficeOpenXml.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotTableAreaStyleCollection : EPPlusReadOnlyList<ExcelPivotTableAreaStyle>
    {
        ExcelStyles _styles;
        XmlHelper _xmlHelper;
        ExcelPivotTable _pt;
        public ExcelPivotTableAreaStyleCollection(ExcelPivotTable pt)
        {
            _pt = pt;
            _styles = pt.WorkSheet.Workbook.Styles;
            foreach (XmlNode node in pt.GetNodes("d:formats/d:format/d:pivotArea"))
            {
                var s = new ExcelPivotTableAreaStyle(_styles.NameSpaceManager, node, _pt);
                _list.Add(s);
            }
        }
        public ExcelPivotTableAreaStyle Add()
        {
            var formatNode = GetTopNode();
            var s = new ExcelPivotTableAreaStyle(_styles.NameSpaceManager, formatNode.FirstChild, _pt);
            _list.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a style for the top right cells of the pivot table, to the right of any filter button, if reading order i set to Left-To-Right. 
        /// </summary>
        /// <returns></returns>
        public ExcelPivotTableAreaStyle AddTopEnd()
        {
            return AddTopEnd(null);
        }
        /// <summary>
        /// Adds a style for the top right cells of the pivot table, to the right of any filter button, if reading order i set to Left-To-Right. 
        /// </summary>
        /// <param name="offsetAddress">Offset address from the top-left cell to the right of any filter button. The top-left cell is refereced as A1. For example, B1:C1 will reference the second and third cell of the first row of the area. "null" will mean all cells</param>
        /// <returns></returns>
        public ExcelPivotTableAreaStyle AddTopEnd(string offsetAddress=null)
        {
            var formatNode = GetTopNode();
            var s = new ExcelPivotTableAreaStyle(_styles.NameSpaceManager, formatNode.FirstChild, _pt)
            {
                PivotAreaType = ePivotAreaType.TopEnd,

            };
            if (offsetAddress != null)
            {
                if(ExcelCellBase.IsSimpleAddress(offsetAddress)==false)
                {
                    throw new ArgumentException("Offset address must be a valid address", "offsetAddress");
                }
                s.Offset = offsetAddress;
            }
            _list.Add(s);
            return s;
        }
        /// <summary>
        /// Adds a style for the top left cells of the pivot table, if reading order i set to Left-To-Right
        /// </summary>
        /// <param name="offsetAddress">Offset address from the left cell. The top-left cell is refereced as A1. For example, B1:C1 will reference the second and third cell of the first row of the area. "null" will mean all cells </param>
        /// <returns></returns>
        public ExcelPivotTableAreaStyle AddTopStart(string offsetAddress = null)
        {
            var formatNode = GetTopNode();
            var s = new ExcelPivotTableAreaStyle(_styles.NameSpaceManager, formatNode.FirstChild, _pt)
            {
                PivotAreaType = ePivotAreaType.Origin,
                FieldIndex = 0,
                FieldPosition = 0,
                LabelOnly = true,
                DataOnly = false
            };
            if (offsetAddress != null)
            {
                if (ExcelCellBase.IsSimpleAddress(offsetAddress) == false)
                {
                    throw new ArgumentException("Offset address must be a valid address", "offsetAddress");
                }
                s.Offset = offsetAddress;
            }
            _list.Add(s);
            return s;
        }
        /// <summary>
        /// Adds a style for the filter box.
        /// </summary>
        /// <param name="field">The field with the box to style</param>
        /// <returns></returns>
        public ExcelPivotTableAreaStyle AddButtonField(ExcelPivotTableField field)
        {
            var formatNode = GetTopNode();
            var s = new ExcelPivotTableAreaStyle(_styles.NameSpaceManager, formatNode.FirstChild, _pt)
            {
                PivotAreaType = ePivotAreaType.FieldButton,
                FieldIndex = field.Index,
                FieldPosition = 0,
                LabelOnly = true,
                DataOnly = false,
                Outline = false
            };

            if (field.IsColumnField)
            {
                s.Axis = ePivotTableAxis.ColumnAxis;
            }
            else if (field.IsRowField)
            {
                s.Axis = ePivotTableAxis.RowAxis;
            }
            else if (field.IsPageField)
            {
                s.Axis = ePivotTableAxis.PageAxis;
            }

            _list.Add(s);
            return s;
        }
        /// <summary>
        /// Adds a style that affects the whole table.
        /// </summary>
        /// <returns>The style object used to set the styles</returns>
        public ExcelPivotTableAreaStyle AddWholeTable()
        {
            return AddAll(false, false);
        }
        public ExcelPivotTableAreaStyle AddAllLabels()
        {
            return AddAll(true, false);
        }
        /// <summary>
        /// Adds a style for labels of a pivot table
        /// </summary>
        /// <param name="fields">The pivot table field that style affects</param>
        /// <returns></returns>
        public ExcelPivotTableAreaStyle AddLabel(params ExcelPivotTableField[] fields)
        {
            var s=Add();
            s.LabelOnly = true;
            s.FieldPosition = 0;
            s.Outline = false;
            foreach (var field in fields)
            {
                s.Conditions.Fields.Add(field);
            }
            return s;
        }
        //public ExcelPivotTableAreaStyle AddData(params ExcelPivotTableField[] fields)
        //{
        //    var s = Add();
        //    s.LabelOnly = false;
        //    s.FieldPosition = 0;
        //    s.Outline = false;
        //    foreach (var field in fields)
        //    {
        //        var r = s.AppliesTo.Add(_pt, field.Index);
        //    }
        //    return s;
        //}

        /// <summary>
        /// Adds a style for the data area of a pivot table
        /// </summary>
        /// <param name="fields"></param>
        /// <returns></returns>
        public ExcelPivotTableAreaStyle AddData(params ExcelPivotTableField[] fields)
        {
            var s = Add();
            s.PivotAreaType = ePivotAreaType.Data;
            s.LabelOnly = false;
            s.FieldPosition = 0;
            s.Outline = false;
            foreach (var field in fields)
            {
                var r = s.Conditions.Fields.Add(_pt, field.Index);
            }
            return s;
        }

        public ExcelPivotTableAreaStyle AddAllData()
        {
            return AddAll(false, true);
        }

        internal ExcelPivotTableAreaStyle AddAll(bool labels, bool data)
        {
            var formatNode = GetTopNode();
            var s = new ExcelPivotTableAreaStyle(_styles.NameSpaceManager, formatNode.FirstChild, _pt)
            {
                PivotAreaType = ePivotAreaType.All,
                LabelOnly = labels,
                DataOnly = data                
            };
            _list.Add(s);
            return s;
        }
        /// <summary>
        /// Adds a style for filter boxes.
        /// </summary>
        /// <param name="axis">The axis for the field buttons</param>
        /// <returns></returns>
        public ExcelPivotTableAreaStyle AddButtonField(ePivotTableAxis axis)
        {
            var formatNode = GetTopNode();
            var s = new ExcelPivotTableAreaStyle(_styles.NameSpaceManager, formatNode.FirstChild, _pt)
            {
                PivotAreaType = ePivotAreaType.FieldButton,
                FieldIndex = 0,
                FieldPosition = 0,
                LabelOnly = true,
                DataOnly = false,
                Outline = false,
                Axis = axis
            };

            _list.Add(s);
            return s;
        }

        internal ExcelPivotTableAreaStyle Add(ePivotAreaType type)
        {
            var formatNode = GetTopNode();
            var s = new ExcelPivotTableAreaStyle(_styles.NameSpaceManager, formatNode.FirstChild, _pt)
            {
                PivotAreaType = type
            };
            _list.Add(s);
            return s;
        }

        internal ExcelPivotTableAreaStyle Add(ePivotAreaType type, ePivotTableAxis axis)
        {
            var formatNode = GetTopNode();
            
            var s = new ExcelPivotTableAreaStyle(_styles.NameSpaceManager, formatNode.FirstChild, _pt)
            {
                PivotAreaType = type,
                Axis = axis
            };
            _list.Add(s);
            return s;
        }
        private XmlNode GetTopNode()
        {
            if (_xmlHelper == null)
            {
                var node = _pt.CreateNode("d:formats");
                _xmlHelper = XmlHelperFactory.Create(_pt.NameSpaceManager, node);
            }
            
            var retNode = _xmlHelper.CreateNode("d:format",false,true);
            retNode.InnerXml = $"<pivotArea xmlns=\"{ExcelPackage.schemaMain}\"/>";
            return retNode;
        }
    }
}
