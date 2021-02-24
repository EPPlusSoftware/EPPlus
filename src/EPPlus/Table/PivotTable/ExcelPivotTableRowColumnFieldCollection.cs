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

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Collection class for row and column fields in a Pivottable 
    /// </summary>
    public class ExcelPivotTableRowColumnFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableField>
    {
        internal string _topNode;
        private readonly ExcelPivotTable _table;

        internal ExcelPivotTableRowColumnFieldCollection(ExcelPivotTable table, string topNode) :
            base()
	    {
            _table = table;
            _topNode=topNode;
	    }

        /// <summary>
        /// Add a new row/column field
        /// </summary>
        /// <param name="Field">The field</param>
        /// <returns>The new field</returns>
        public ExcelPivotTableField Add(ExcelPivotTableField Field)
        {
            if(Field==null)
            {
                throw (new ArgumentNullException("Field","Pivot Table Field can't be null"));
            }
            if((_topNode=="colFields" && Field.DragToCol==false))
            {
                throw (new ArgumentException("Field", "This field is not allowed as a column field."));
            }

            if ((_topNode == "rowFields" && Field.DragToRow == false))
            {
                throw (new ArgumentException("Field", "This field is not allowed as a row field."));
            }

            if ((_topNode == "pageFields" && Field.DragToPage == false))
            {
                throw (new ArgumentException("Field", "This field is not allowed as a Page field."));
            }


            SetFlag(Field, true);
            _list.Add(Field);
            return Field;
        }
        /// <summary>
        /// Insert a new row/column field
        /// </summary>
        /// <param name="Field">The field</param>
        /// <param name="Index">The position to insert the field</param>
        /// <returns>The new field</returns>
        internal ExcelPivotTableField Insert(ExcelPivotTableField Field, int Index)
        {
            SetFlag(Field, true);
            _list.Insert(Index, Field);
            return Field;
        }
        private void SetFlag(ExcelPivotTableField field, bool value)
        {
            switch (_topNode)
            {
                case "rowFields":
                    if (field.IsColumnField || field.IsPageField)
                    {
                        throw(new Exception("This field is a column or page field. Can't add it to the RowFields collection"));
                    }
                    field.IsRowField = value;
                    field.Axis = ePivotFieldAxis.Row;
                    break;
                case "colFields":
                    if (field.IsRowField || field.IsPageField)
                    {
                        throw (new Exception("This field is a row or page field. Can't add it to the ColumnFields collection"));
                    }
                    field.IsColumnField = value;
                    field.Axis = ePivotFieldAxis.Column;
                    break;
                case "pageFields":
                    if (field.IsColumnField || field.IsRowField)
                    {
                        throw (new Exception("Field is a column or row field. Can't add it to the PageFields collection"));
                    }
                    if (_table.Address._fromRow < 3)
                    {
                        throw(new Exception(string.Format("A pivot table with page fields must be located above row 3. Currenct location is {0}", _table.Address.Address)));
                    }
                    field.IsPageField = value;
                    field.Axis = ePivotFieldAxis.Page;
                    break;
                case "dataFields":
                    
                    break;
            }
        }
        /// <summary>
        /// Remove a field
        /// </summary>
        /// <param name="Field"></param>
        public void Remove(ExcelPivotTableField Field)
        {
            if(!_list.Contains(Field))
            {
                throw new ArgumentException("Field not in collection");
            }
            SetFlag(Field, false);            
            _list.Remove(Field);            
        }
        /// <summary>
        /// Remove a field at a specific position
        /// </summary>
        /// <param name="Index"></param>
        public void RemoveAt(int Index)
        {
            if (Index > -1 && Index < _list.Count)
            {
                throw(new IndexOutOfRangeException());
            }
            SetFlag(_list[Index], false);
            _list.RemoveAt(Index);      
        }
    }
}