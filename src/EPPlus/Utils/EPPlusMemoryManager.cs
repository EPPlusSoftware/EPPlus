using System;
/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/
  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/27/2025         EPPlus Software AB       EPPlus 7.7
 *************************************************************************************************/
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils
{
    /// <summary>
    /// The purpose of this class is to encapsulate the calls the <see cref="RecyclableMemory"/>'s GetStream methods.
    /// By wrapping these methods we can catch errors like <see cref="FileNotFoundException"/> and <see cref="FileLoadException"/>.
    /// </summary>
    internal static class EPPlusMemoryManager
    {
        // this variable indicates if an error has occured while calling
        // RecyclableMemory
        private static bool _isBroken = false;

        // indicates if RecyclableMemory should be used
        public static bool UseRecyclableMemory
        {
            get; set;
        } = false;

        private static void Log(Exception e, string message)
        {
            // implement logging here
        }

        public static MemoryStream GetStream()
        {
            if (_isBroken || !UseRecyclableMemory)
                return new MemoryStream();

            try
            {
                return RecyclableMemory.GetStreamInternal();
            }
            catch (Exception e)
            {
                Log(e, "Failed to get stream from RecyclableMemory");
                _isBroken = true;
                return new MemoryStream();
            }
        }

        public static MemoryStream GetStream(byte[] buffer)
        {
            if (_isBroken || !UseRecyclableMemory)
                return new MemoryStream(buffer);

            try
            {
                return RecyclableMemory.GetStreamInternal(buffer);
            }
            catch (Exception e)
            {
                Log(e, "Failed to get stream from RecyclableMemory");
                _isBroken = true;
                return new MemoryStream(buffer);
            }
        }

        public static MemoryStream GetStream(int capacity)
        {
            if (_isBroken || !UseRecyclableMemory)
                return new MemoryStream(capacity);

            try
            {
                return RecyclableMemory.GetStreamInternal(capacity);
            }
            catch (Exception e)
            {
                Log(e, "Failed to get stream from RecyclableMemory");
                _isBroken = true;
                return new MemoryStream(capacity);
            }
        }

    }
}