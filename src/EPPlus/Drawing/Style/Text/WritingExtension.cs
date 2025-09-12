using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.WritingExtension
{
    internal static class WritingExtension
    {
        /// <summary>
        /// Act if not null
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="param"></param>
        /// <param name="method"></param>
        internal static bool TryAct<T>(this T param, Action<T> method)
        {
            if (param != null)
            {
                method(param); 
                return true;
            }
            return false;
        }
    }
}

