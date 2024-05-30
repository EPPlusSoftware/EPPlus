/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/07/2023         EPPlus Software AB         Implemented function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Operators
{
    /// <summary>
    /// Implements the KahanSum algorithm to reduce floating point errors.
    /// </summary>
    internal class KahanSum
    {
        public KahanSum(double d)
        {
            Add(d);
        }
        private double _sum;
        private double _c;

        public KahanSum Add(double d)
        {
            var t = _sum + d;
            if (_sum >= d)
            {
                _c += (_sum - t) + d; // If sum is bigger, low - order digits of input[i] are lost.
            }
            else
            {
                _c += (d - t) + _sum; // Else low-order digits of sum are lost.
            }
            _sum = t;
            return this;
        }

        public static KahanSum operator +(KahanSum a, double b)
        {
            return a.Add(b);
        }

        public static KahanSum operator +(KahanSum a, KahanSum b)
        {
            return a.Add(b.Get());
        }

        public static implicit operator KahanSum(double d) => new KahanSum(d);

        public static implicit operator double(KahanSum kh) => kh.Get();

        public double Get()
        {
            return _sum + _c;
        }

        public void Clear()
        {
            _sum = 0.0;
            _c = 0.0;
        }
    }
}
