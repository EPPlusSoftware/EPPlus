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
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A list of pivot item refernces
    /// </summary>
    public class ExcelPivotAreaReferenceItems : EPPlusReadOnlyList<PivotItemReference>
    {
        private ExcelPivotAreaReference _reference;
        internal ExcelPivotAreaReferenceItems(ExcelPivotAreaReference reference)
        {
            _reference = reference;
        }
        /// <summary>
        /// Adds the item at the index to the condition. The index referes to the pivot cache.
        /// </summary>
        /// <param name="index">Index into the pivot cache items. Either the shared items or the group items</param>
        public void Add(int index)
        {
            {
                var items = _reference.Field.Cache.SharedItems.Count == 0 ? _reference.Field.Cache.GroupItems : _reference.Field.Cache.SharedItems;
                if (items.Count > index)
                {
                    Add(new PivotItemReference() { Index = index, Value = items[index] });
                }
                else
                {
                    throw new IndexOutOfRangeException("Index is out of range in cache Items. Please make sure the pivot table cache has been refreshed.");
                }
            }
        }
        /// <summary>
        /// Adds a specific cache item to the condition. The value is matched against the values in the pivot cache, either the shared items or the group items.
        /// </summary>
        /// <param name="value">The value to match against. Is matched agaist the cache values and must be matched with the same data type.</param>
        /// <returns>true if the value has been added, otherwise false</returns>
        public bool AddByValue(object value)
        {
            var index = _reference.Field.Items._list.FindIndex(x => (x.Value!=null && (x.Value.Equals(value)) || (x.Text!=null && x.Text.Equals(value))));
            if (index >= 0)
            {
                Add(new PivotItemReference() { Index = index, Value = value });
                return true;
            }
            return false;
        }
    }
}