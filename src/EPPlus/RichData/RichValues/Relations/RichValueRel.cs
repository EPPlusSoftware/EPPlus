﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues.Relations
{
    internal class RichValueRel
    {
        public string Id { get; set; }
        public string Type { get; set; }

        public Uri TargetUri { get; set; }

        internal void WriteXml(StreamWriter sw)
        {
            sw.Write($"<rel r:id=\"{Id}\" />");
        }
    }
}