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
        /// <summary>
        /// Sets the show data as to type Normal. This removes the Show data as setting.
        /// </summary>
        public void SetNormal()
        {
            _dataField.ShowDataAsInternal = eShowDataAs.Normal;
            _dataField.BaseField = 0;
            _dataField.BaseItem = 0;
        }

        /// <summary>
        /// Sets the show data as to type Percent Of Total
        /// </summary>
        public void SetPercentOfTotal()
        {
            _dataField.ShowDataAsInternal = eShowDataAs.PercentOfTotal;
            _dataField.BaseField = 0;
            _dataField.BaseItem = 0;
        }
        /// <summary>
        /// Sets the show data as to type Percent Of Row
        /// </summary>
        public void SetPercentOfRow()
        {
            _dataField.ShowDataAsInternal = eShowDataAs.PercentOfRow;
            _dataField.BaseField = 0;
            _dataField.BaseItem = 0;
        }
        /// <summary>
        /// Sets the show data as to type Percent Of Column
        /// </summary>
        public void SetPercentOfColumn()
        {
            _dataField.ShowDataAsInternal = eShowDataAs.PercentOfColumn;
            _dataField.BaseField = 0;
            _dataField.BaseItem = 0;
        }
        /// <summary>
        /// Sets the show data as to type Percent
        /// <param name="baseField">The base field to use</param>
        /// <param name="baseItem">The index of the item to use within the <see cref="ExcelPivotTableField.Items"/> collection of the base field</param>
        /// </summary>
        public void SetPercent(ExcelPivotTableField baseField, int baseItem)
        {
            Validate(baseField, baseItem);
            _dataField.ShowDataAsInternal = eShowDataAs.Percent;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = baseItem;
        }
        /// <summary>
        /// Sets the show data as to type Percent
        /// <param name="baseField">The base field to use</param>
        /// <param name="baseItem">The previous or next field</param>
        /// </summary>
        public void SetPercent(ExcelPivotTableField baseField, ePrevNextPivotItem baseItem)
        {
            Validate(baseField);
            _dataField.ShowDataAsInternal = eShowDataAs.Percent;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = (int)baseItem;
        }

        /// <summary>
        /// Sets the show data as to type Percent Of Parent
        /// <param name="baseField">The base field to use</param>
        /// </summary>
        public void SetPercentParent(ExcelPivotTableField baseField)
        {
            Validate(baseField);
            _dataField.ShowDataAsInternal = eShowDataAs.PercentOfParent;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = 0;
        }

        /// <summary>
        /// Sets the show data as to type Index
        /// </summary>
        public void SetIndex()
        {
            _dataField.ShowDataAsInternal = eShowDataAs.Index;
            _dataField.BaseField = 0;
            _dataField.BaseItem = 0;
        }

        /// <summary>
        /// Sets the show data as to type Running Total
        /// </summary>
        /// <param name="baseField">The base field to use</param>
        public void SetRunningTotal(ExcelPivotTableField baseField)
        {
            Validate(baseField);
            _dataField.ShowDataAsInternal = eShowDataAs.RunningTotal;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = 0;
        }
        /// <summary>
        /// Sets the show data as to type Difference
        /// </summary>
        /// <param name="baseField">The base field to use</param>
        /// <param name="baseItem">The index of the item to use within the <see cref="ExcelPivotTableField.Items"/> collection of the base field</param>
        public void SetDifference(ExcelPivotTableField baseField, int baseItem)
        {
            Validate(baseField, baseItem);
            _dataField.ShowDataAsInternal = eShowDataAs.Difference;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = baseItem;
        }
        /// <summary>
        /// Sets the show data as to type Difference
        /// </summary>
        /// <param name="baseField">The base field to use</param>
        /// <param name="baseItem">The previous or next field</param>
        public void SetDifference(ExcelPivotTableField baseField, ePrevNextPivotItem baseItem)
        {
            Validate(baseField);
            _dataField.ShowDataAsInternal = eShowDataAs.Difference;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = (int)baseItem;
        }

        /// <summary>
        /// Sets the show data as to type Percent Of Total
        /// <param name="baseField">The base field to use</param>
        /// <param name="baseItem">The index of the item to use within the <see cref="ExcelPivotTableField.Items"/> collection of the base field</param>
        /// </summary>
        public void SetPercentageDifference(ExcelPivotTableField baseField, int baseItem)
        {
            Validate(baseField);
            _dataField.ShowDataAsInternal = eShowDataAs.PercentDifference;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = baseItem;
        }
        /// <summary>
        /// Sets the show data as to type Percent Of Total
        /// <param name="baseField">The base field to use</param>
        /// <param name="baseItem">The previous or next field</param>
        /// </summary>
        public void SetPercentageDifference(ExcelPivotTableField baseField, ePrevNextPivotItem baseItem)
        {
            Validate(baseField);
            _dataField.ShowDataAsInternal = eShowDataAs.PercentDifference;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = (int)baseItem;
        }

        /// <summary>
        /// Sets the show data as to type Percent Of Parent Row
        /// </summary>
        public void SetPercentParentRow()
        {
            _dataField.ShowDataAsInternal = eShowDataAs.PercentOfParentRow;
        }
        /// <summary>
        /// Sets the show data as to type Percent Of Parent Column
        /// </summary>
        public void SetPercentParentColumn()
        {
            _dataField.ShowDataAsInternal = eShowDataAs.PercentOfParentColumn;
        }
        /// <summary>
        /// Sets the show data as to type Percent Of Running Total
        /// </summary>
        public void SetPercentOfRunningTotal(ExcelPivotTableField baseField)
        {
            Validate(baseField);
            _dataField.ShowDataAsInternal = eShowDataAs.PercentOfRunningTotal;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = 0;
        }
        /// <summary>
        /// Sets the show data as to type Rank Ascending
        /// <param name="baseField">The base field to use</param>
        /// </summary>
        public void SetRankAscending(ExcelPivotTableField baseField)
        {
            Validate(baseField);
            _dataField.ShowDataAsInternal = eShowDataAs.RankAscending;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = 0;
        }
        /// <summary>
        /// Sets the show data as to type Rank Descending
        /// <param name="baseField">The base field to use</param>
        /// </summary>
        public void SetRankDescending(ExcelPivotTableField baseField)
        {
            Validate(baseField);
            _dataField.ShowDataAsInternal = eShowDataAs.RankDescending;
            _dataField.BaseField = baseField.Index;
            _dataField.BaseItem = 0;
        }
        /// <summary>
        /// The value of the "Show Data As" setting
        /// </summary>
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