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

namespace OfficeOpenXml.Core.CellStore
{
    /// <summary>
    /// These binary search functions are identical, exept that one uses a struc and the other a class.
    /// Structs consume less memory and are also faster.
    /// </summary>
    internal static class ArrayUtil
    {
        /// <summary>
        /// For the struct.
        /// </summary>
        /// <param name="store"></param>
        /// <param name="pos"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        internal static int OptimizedBinarySearch(IndexItem[] store, int pos, int length)
        {
                if (length == 0) return -1;
                int low = 0, high = length - 1, mid;

                while (low <= high)
                {
                    mid = (low + high) >> 1;

                    if (pos < store[mid].Index)
                        high = mid - 1;

                    else if (pos > store[mid].Index)
                        low = mid + 1;

                    else
                        return mid;
                }
                return ~low;
        }
        internal static int OptimizedBinarySearch(IndexBase[] store, int pos, int length)
        {
            if (length == 0) return -1;
            int low = 0, high = length - 1, mid;

            while (low <= high)
            {
                mid = (low + high) >> 1;

                if (pos < store[mid].Index)
                    high = mid - 1;

                else if (pos > store[mid].Index)
                    low = mid + 1;

                else
                    return mid;
            }
            return ~low;
        }
    }
}