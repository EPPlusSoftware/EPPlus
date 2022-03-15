using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Core
{
    /// <summary>
    /// A readonly collection of a generic type
    /// </summary>
    /// <typeparam name="T">The generic type</typeparam>
    public class EPPlusReadOnlyList<T> : IEnumerable<T>
    {
        internal List<T> _list=new List<T>();
        IEnumerator<T> IEnumerable<T>.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        /// <summary>
        /// Return the enumerator for the collection
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        /// <summary>
        /// The indexer for the collection
        /// </summary>
        /// <param name="index">The index</param>
        /// <returns>Returns the object at the index</returns>
        public T this[int index]
        {
            get
            {
                return _list[index];
            }
        }
        ///// <summary>
        ///// Gets the item with the value supplied
        ///// </summary>
        ///// <param name="value">The values</param>
        ///// <returns>The </returns>
        //public T GetByValue(T value)
        //{
        //    var ix=_list.IndexOf(value);
        //    if(ix<0)
        //    {
        //        return _list[ix];
        //    }
        //    return default;
        //}
        /// <summary>
        /// Retrives the index of the supplied value
        /// </summary>
        /// <param name="value"></param>
        /// <returns>The index</returns>
        public int GetIndexByValue(T value)
        {
            return _list.IndexOf(value);
        }
        public int Count
        {
            get
            {
                return _list.Count;
            }
        }
        internal virtual void Clear()
        {
            _list.Clear();
        }
        internal virtual void Add(T item)
        {
            _list.Add(item);
        }

    }
}
