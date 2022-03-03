/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/11/2021         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Accessibility
{
    internal class AriaAttributes
    {
        internal class AriaAttribute
        {
            public AriaAttribute(string attributeName, string defaultValue)
            {
                AttributeName = attributeName;
                DefaultValue = defaultValue;
            }

            public string AttributeName { get; private set; }

            public string DefaultValue { get; private set; }
        }

        public static AriaAttribute AriaLabel { get { return new AriaAttribute("aria-label", string.Empty); } }

        public static AriaAttribute AriaDescribedBy { get { return new AriaAttribute("aria-describedby", string.Empty); } }

        public static AriaAttribute AriaLabelledBy { get { return new AriaAttribute("aria-labelledby", string.Empty); } }
    }
}
