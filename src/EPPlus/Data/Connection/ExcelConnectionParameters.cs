/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// A collection of connection parameters.
    /// </summary>
    public class ExcelConnectionParameters : IEnumerable<ExcelConnectionParameter>
    {
        List<ExcelConnectionParameter> _list = new List<ExcelConnectionParameter>();

        /// <summary>
        /// Indexer for the collection
        /// </summary>
        /// <param name="index">The index of the parameter to get.</param>
        /// <returns>The parameter.</returns>
        public ExcelConnectionParameter this[int index]
        {
            get
            {
                return _list[index];
            }
        }
        /// <summary>
        /// Number of items in the collection.
        /// </summary>
        public int Count { get { return _list.Count; } }
        /// <summary>
        /// The enumerator
        /// </summary>
        /// <returns></returns>
        public IEnumerator<ExcelConnectionParameter> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        internal void Add(ExcelConnectionParameter parameter)
        {
            _list.Add(parameter);
        }   
        /// <summary>
        /// Adds a new blank parameter to the collection and returns it.
        /// </summary>
        /// <returns></returns>
        public ExcelConnectionParameter Add()
        {
            var para= new ExcelConnectionParameter();
            _list.Add(para);
            return para;
        }
        /// <summary>
        /// Remove the parameter at the index supplied
        /// </summary>
        /// <param name="index">The index of the parameter to remove.</param>
        public void RemoveAt(int index)
        {
            _list.RemoveAt(index);
        }
        /// <summary>
        /// Removes the parameter from the collection.
        /// </summary>
        /// <param name="parameter"></param>
        public void Remove(ExcelConnectionParameter parameter)
        {
            _list.Remove(parameter);
        }
    }
}