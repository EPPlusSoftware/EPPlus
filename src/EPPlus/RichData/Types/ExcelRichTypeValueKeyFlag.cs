/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       EPPlus 7.4
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.RichData.Types
{
    internal class ExcelRichTypeValueKeyFlag
    {
        public ExcelRichTypeValueKeyFlag(RichValueKeyFlags flag, bool value)
        {
            Flag = flag;
            Value = value;
        }

        public RichValueKeyFlags Flag { get; }
        public bool Value { get; }


        public void WriteXml(StreamWriter sr)
        {
            var val = Value ? "1" : "0";
            sr.Write($"<flag name=\"{Flag.ToString()}\" value=\"{val}\" />");
        }
    }
}
