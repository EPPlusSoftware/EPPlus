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
    /// Compares the item to the previous or next item.
    /// </summary>
    public enum ePrevNextPivotItem
    {
        /// <summary>
        /// Previous item
        /// </summary>
        Previous = 1048828,
        /// <summary>
        /// Next item
        /// </summary>
        Next = 1048829
    }

    public class ExcelPivotTableDataFieldShowDataAs
    {
        ExcelPivotTableDataField _dataField;
        public ExcelPivotTableDataFieldShowDataAs(ExcelPivotTableDataField dataField)
        {
            _dataField = dataField;
        }
        public void SetNormal()
        {
            _dataField.ShowDataAsInternal = eShowDataAs.Normal;
            _dataField.BaseField = 0;
            _dataField.BaseItem = 0;
        }

        public void SetPercentOfTotal()
        {
            _dataField.ShowDataAsInternal = eShowDataAs.PercentOfTotal;
            _dataField.BaseField = 0;
            _dataField.BaseItem = 0;
        }
        public void SetPercentOfRow()
        {
            _dataField.ShowDataAsInternal = eShowDataAs.PercentOfRow;
            _dataField.BaseField = 0;
            _dataField.BaseItem = 0;
        }
        public void SetPercentOfColumn()
        {
            _dataField.ShowDataAsInternal = eShowDataAs.PercentOfCol;
            _dataField.BaseField = 0;
            _dataField.BaseItem = 0;
        }        
        public void SetPercent(ExcelPivotTableField baseField, int baseItem)
        {
            Validate(baseField, baseItem);
            _dataField.ShowDataAsInternal = eShowDataAs.Percent;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = baseItem;
        }
        public void SetPercent(ExcelPivotTableField baseField, ePrevNextPivotItem baseItem)
        {
            Validate(baseField);
            _dataField.ShowDataAsInternal = eShowDataAs.Percent;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = (int)baseItem;
        }

        public void SetPercentParent(ExcelPivotTableField baseField)
        {
            Validate(baseField);
            _dataField.ShowDataAsInternal = eShowDataAs.PercentOfParent;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = 0;
        }

        public void SetIndex()
        {
            _dataField.ShowDataAsInternal = eShowDataAs.Index;
            _dataField.BaseField = 0;
            _dataField.BaseItem = 0;
        }

        public void SetRunningTotal(ExcelPivotTableField baseField)
        {
            Validate(baseField);
            _dataField.ShowDataAsInternal = eShowDataAs.RunTotal;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = 0;
        }
        public void SetDifference(ExcelPivotTableField baseField, int baseItem)
        {
            Validate(baseField, baseItem);
            _dataField.ShowDataAsInternal = eShowDataAs.Difference;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = baseItem;
        }
        public void SetDifference(ExcelPivotTableField baseField, ePrevNextPivotItem baseItem)
        {
            Validate(baseField);
            _dataField.ShowDataAsInternal = eShowDataAs.Difference;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = (int)baseItem;
        }

        public void SetPercentageDifference(ExcelPivotTableField baseField, ePrevNextPivotItem baseItem)
        {
            Validate(baseField);
            _dataField.ShowDataAsInternal = eShowDataAs.PercentDiff;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = (int)baseItem;
        }

        public void SetPercentParentRow()
        {
            _dataField.ShowDataAsInternal = eShowDataAs.PercentOfParentRow;
        }
        public void SetPercentParentColumn()
        {
            _dataField.ShowDataAsInternal = eShowDataAs.PercentOfParentCol;
        }

        public eShowDataAs Value
        {
            get
            {
                return _dataField.ShowDataAsInternal;
            }
        }
        private void Validate(ExcelPivotTableField baseField, int? baseItem = null)
        {
            if (baseField._pivotTable != _dataField.Field._pivotTable)
            {
                throw new ArgumentException("The base field must be from the same pivot table as the data field", nameof(baseField));
            }
            if (baseField == _dataField.Field)
            {
                throw new ArgumentException("The base field and the data field must not be the same.", nameof(baseField));
            }
            if (baseItem != null)
            {
                if (baseItem<0 || baseItem >= baseField.Items.Count)
                {
                    throw new ArgumentException("Base items must be within an index the fields item collection. Please refresh the Items collection of the field to get the items from source.", nameof(baseField));
                }
            }
        }

    }
}