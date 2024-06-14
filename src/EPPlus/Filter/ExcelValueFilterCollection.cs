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
namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// A collection of value filters
    /// </summary>
    public class ExcelValueFilterCollection : ExcelFilterCollectionBase<ExcelFilterItem>
    {
        /// <summary>
        /// Flag indicating whether to filter by blank
        /// </summary>
        public bool Blank { get; set; }
        /// <summary>
        /// The calendar to be used. To be implemented
        /// </summary>
        internal eCalendarType? CalendarType{get;set;}
        /// <summary>
        /// Add a Date filter item. 
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public ExcelFilterDateGroupItem Add(ExcelFilterDateGroupItem value)
        {
            _list.Add(value);
            return value;
        }
        /// <summary>
        /// <para>Add a filter value that will be matched agains the ExcelRange.String property</para>
        /// If value is "" or null sets Blank=True instead of adding.
        /// </summary>
        /// <param name="item">The value to add. If "" or null sets Blank=True instead.</param>
        /// <returns>The filter value item</returns>
        public ExcelFilterValueItem Add(ExcelFilterValueItem item)
        {
            AddOrSetBlank(item);
            return item;
        }
        /// <summary>
        /// <para>Add a filter value that will be matched agains the ExcelRange.Text property</para>
        /// If value is "" or null sets Blank=True instead of adding.
        /// </summary>
        /// <param name="value">The value to add. If "" or null sets Blank=True instead.</param>
        /// <returns>The filter value item</returns>
        public ExcelFilterValueItem Add(string value)
        {
            var v = new ExcelFilterValueItem(value);
            AddOrSetBlank(v);
            return v;
        }

        internal void AddOrSetBlank(ExcelFilterValueItem item)
        {
            if (string.IsNullOrEmpty(item.Value))
            {
                Blank = true;
            }
            else
            {
                _list.Add(item);
            }
        }

        /// <summary>
        /// Clears the collection
        /// </summary>
        public void Clear()
        {
            _list.Clear();
        }
        /// <summary>
        /// Remove the item at the specified index from the list
        /// </summary>
        /// <param name="index">The index in the list</param>
        public void RemoveAt(int index)
        {
            _list.RemoveAt(index);
        }
        /// <summary>
        /// Remove the item from the list
        /// </summary>
        /// <param name="item">The item to remove</param>
        public void Remove(ExcelFilterItem item)
        {
            _list.Remove(item);
        }
    }
}