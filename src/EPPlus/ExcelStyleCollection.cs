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
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using System.Linq;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml
{
    public class ExcelNamedStyleCollection<T> : ExcelStyleCollection<T>
    {
        /// <summary>
        /// Indexer for the collection
        /// </summary>
        /// <param name="name">The name of the Style</param>
        /// <returns></returns>
        public T this[string name]
        {
            get
            {
                if(_dic.ContainsKey(name))
                {
                    return _list[_dic[name]];
                }
                return default(T);
            }
        }
    }
    /// <summary>
    /// Base collection class for styles.
    /// </summary>
    /// <typeparam name="T">The style type</typeparam>
    public class ExcelStyleCollection<T> : IEnumerable<T>
    {
        internal ExcelStyleCollection()
        {
            _setNextIdManual = false;
        }
        bool _setNextIdManual;
        internal ExcelStyleCollection(bool SetNextIdManual)
        {
            _setNextIdManual = SetNextIdManual;
        }
        /// <summary>
        /// The top xml node of the collection
        /// </summary>
        public XmlNode TopNode { get; set; }
        internal List<T> _list = new List<T>();
        protected internal Dictionary<string, int> _dic = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        internal int NextId=0;
        #region IEnumerable<T> Members
        /// <summary>
        /// Returns an enumerator that iterates through a collection.
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<T> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        #endregion
        #region IEnumerable Members
        /// <summary>
        /// Returns an enumerator that iterates through a collection.
        /// </summary>
        /// <returns>The enumerator</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        #endregion
        /// <summary>
        /// Indexer for the collection
        /// </summary>
        /// <param name="PositionID">The index of the Style</param>
        /// <returns></returns>
        public T this[int PositionID]
        {
            get
            {
                return _list[PositionID];
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
        internal int Add(string key, T item)
        {
            _list.Add(item);
            if (!_dic.ContainsKey(key.ToLower(CultureInfo.InvariantCulture))) _dic.Add(key.ToLower(CultureInfo.InvariantCulture), _list.Count - 1);
            if (_setNextIdManual) NextId++;
            return _list.Count-1;
        }
        /// <summary>
        /// Finds the key 
        /// </summary>
        /// <param name="key">the key to be found</param>
        /// <param name="obj">The found object.</param>
        /// <returns>True if found</returns>
        internal bool FindById(string key, ref T obj)
        {
            if (_dic.ContainsKey(key))
            {
                obj = _list[_dic[key.ToLower(CultureInfo.InvariantCulture)]];
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// Find Index
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        internal int FindIndexById(string key)
        {
            if (_dic.ContainsKey(key))
            {
                return _dic[key];
            }
            else
            {
                return int.MinValue;
            }
        }
        internal int FindIndexByBuildInId(int id)
        {
            for(int i=0;i<_list.Count;i++)
            {
                if (_list[i] is ExcelNamedStyleXml ns)
                {
                    if (ns.BuildInId == id)
                    {
                        return i;
                    }
                }
            }
            return -1;
        }

        internal bool ExistsKey(string key)
        {
            return _dic.ContainsKey(key);
        }
        internal void Sort(Comparison<T> c)
        {
            _list.Sort(c);
        }
    }
}
