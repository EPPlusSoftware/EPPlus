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
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
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

        private static bool _isBroken = false;

        public static MemoryStream GetStream()
        {
            if (_isBroken)
                return new MemoryStream();

            try
            {
                return RecyclableMemory.GetStreamInternal();
            }
            catch (Exception)
            {
                _isBroken = true;
                return new MemoryStream();
            }
        }

        public static MemoryStream GetStream(byte[] buffer)
        {
            if (_isBroken)
                return new MemoryStream(buffer);

            try
            {
                return RecyclableMemory.GetStreamInternal(buffer);
            }
            catch (Exception)
            {
                _isBroken = true;
                return new MemoryStream(buffer);
            }
        }

        public static MemoryStream GetStream(int capacity)
        {
            if (_isBroken)
                return new MemoryStream(capacity);

            try
            {
                return RecyclableMemory.GetStreamInternal(capacity);
            }
            catch (Exception)
            {
                _isBroken = true;
                return new MemoryStream(capacity);
            }
        }

    }
}
