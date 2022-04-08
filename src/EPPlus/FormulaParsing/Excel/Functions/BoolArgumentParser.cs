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
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class BoolArgumentParser : ArgumentParser
    {
        public override object Parse(object obj)
        {
            if (obj is IRangeInfo)
            {
                var r = ((IRangeInfo)obj).FirstOrDefault();
                obj = (r == null ? null : r.Value);
            }
            if (obj == null) return false;
            if (obj is bool) return (bool)obj;
            if (obj.IsNumeric()) return Convert.ToBoolean(obj);
            bool result;
            if (bool.TryParse(obj.ToString(), out result))
            {
                return result;
            }
            return result;
        }

        public override object Parse(object obj, RoundingMethod roundingMethod)
        {
            return Parse(obj);
        }
    }
}
