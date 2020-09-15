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
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// A collection of filters for a filter column
    /// </summary>
    /// <typeparam name="T">The filter type</typeparam>
    public class ExcelFilterCollectionBase<T> : IEnumerable<T>
    {
        /// <summary>
        /// A list of columns
        /// </summary>
        protected internal List<T> _list;
        protected internal readonly bool _maxTwoItems;
        internal ExcelFilterCollectionBase()
        {
            if (typeof(T) == typeof(ExcelFilterCustomItem))
            {
                _maxTwoItems = true;
            }
            _list = new List<T>();
        }
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<T> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        public T this[int index]
        {
            get
            {
                return _list[index];
            }
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _list.Count;
            }
        }
    }
    /// <summary>
    /// A collection of filters for a filter column
    /// </summary>
    /// <typeparam name="T">The filter type</typeparam>
    public class ExcelFilterCollection<T> : ExcelFilterCollectionBase<T>
    {
        /// <summary>
        /// Add a new filter item
        /// </summary>
        /// <param name="value"></param>
        public T Add(T value)
        {
            if (_maxTwoItems && _list.Count >= 2)
            {
                throw (new InvalidOperationException("You can only have two filters on an ExcelCustomFilterColumn collection"));
            }
            _list.Add(value);
            return value;
        }

    }
}