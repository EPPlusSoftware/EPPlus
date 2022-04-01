/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/16/2020         EPPlus Software AB       EPPlus 5.2.1
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.LoadFunctions.Params
{
    /// <summary>
    /// Parameters for the LoadFromDictionaries method
    /// </summary>
    public class LoadFromTextParams
    {
        /// <summary>
        /// The first row in the text is the header row
        /// </summary>
        public bool FirstRowIsHeader { get; set; }
        /// <summary>
        /// The text to split
        /// </summary>
        public string Text { get; set; }
        /// <summary>
        /// Describes how to split a CSV text.
        /// </summary>
        public ExcelTextFormat Format { get; set; } = new ExcelTextFormat();
    }
}
