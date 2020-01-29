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
using OfficeOpenXml.Compatibility;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.VBA
{
    /// <summary>
    /// Base class for VBA collections
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExcelVBACollectionBase<T> : IEnumerable<T>
    {
        /// <summary>
        /// A list of vba objects
        /// </summary>
        internal protected List<T> _list=new List<T>();
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<T> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        /// <summary>
        /// Indexer
        /// </summary>
        /// <param name="Name">Name</param>
        /// <returns></returns>
        public T this [string Name]
        {
            get
            {
                return _list.Find((f) => TypeCompat.GetPropertyValue(f,"Name").ToString().Equals(Name,StringComparison.OrdinalIgnoreCase));
            }
        }
        /// <summary>
        /// Indexer
        /// </summary>
        /// <param name="Index">Position</param>
        /// <returns></returns>
        public T this[int Index]
        {
            get
            {
                return _list[Index];
            }
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get { return _list.Count; }
        }
        /// <summary>
        /// If a specific name exists in the collection
        /// </summary>
        /// <param name="Name">The name</param>
        /// <returns>True if the name exists</returns>
        public bool Exists(string Name)
        {
            return _list.Exists((f) => TypeCompat.GetPropertyValue(f,"Name").ToString().Equals(Name,StringComparison.OrdinalIgnoreCase));
        }
        /// <summary>
        /// Removes the item
        /// </summary>
        /// <param name="Item"></param>
        public void Remove(T Item)
        {
            _list.Remove(Item);
        }
        /// <summary>
        /// Removes the item at the specified index
        /// </summary>
        /// <param name="index">THe index</param>
        public void RemoveAt(int index)
        {
            _list.RemoveAt(index);
        }
        
        internal void Clear()
        {
            _list.Clear();
        }
    }
}