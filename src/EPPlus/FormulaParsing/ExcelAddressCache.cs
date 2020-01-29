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
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Caches string by generated id's.
    /// </summary>
    public class ExcelAddressCache
    {
        private readonly object _myLock = new object();
        private readonly Dictionary<int, string> _addressCache = new Dictionary<int, string>();
        private readonly Dictionary<string, int> _lookupCache = new Dictionary<string, int>();
        private int _nextId = 1;
        private const bool EnableLookupCache = false;

        /// <summary>
        /// Returns an id to use for caching (when the <see cref="Add"/> method is called)
        /// </summary>
        /// <returns></returns>
        public int GetNewId()
        {
            lock(_myLock)
            {
                return _nextId++;
            }
        }

        /// <summary>
        /// Adds an address to the cache
        /// </summary>
        /// <param name="id"></param>
        /// <param name="address"></param>
        /// <returns></returns>
        public bool Add(int id, string address)
        {
            lock(_myLock)
            {
                if (_addressCache.ContainsKey(id)) return false;
                _addressCache.Add(id, address);
                if(EnableLookupCache && !_lookupCache.ContainsKey(address))
                    _lookupCache.Add(address, id);
                return true;
            }
            
        }

        /// <summary>
        /// Number of items in the cache
        /// </summary>
        public int Count
        {
            get { return _addressCache.Count; }
        }

        /// <summary>
        /// Returns an address by its cache id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public string Get(int id)
        {
            if (!_addressCache.ContainsKey(id)) return string.Empty;
            return _addressCache[id];
        }

        /// <summary>
        /// Clears the cache
        /// </summary>
        public void Clear()
        {
            lock(_myLock)
            {
                _addressCache.Clear();
                _lookupCache.Clear();
                _nextId = 1;
            }  
        }

    }
}
