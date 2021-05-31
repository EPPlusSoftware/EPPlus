/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/07/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sorting
{
    /// <summary>
    /// Base class for Sort options.
    /// </summary>
    public abstract class SortOptionsBase
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public SortOptionsBase()
        {
            ColumnIndexes = new List<int>();
            RowIndexes = new List<int>();
            Descending = new List<bool>();
            CustomLists = new Dictionary<int, string[]>();
            CompareOptions = CompareOptions.None;
        }

        internal bool LeftToRight { get; set; }

        internal List<int> ColumnIndexes { get; private set; }

        internal List<int> RowIndexes { get; private set; }
        internal List<bool> Descending { get; private set; }

        internal Dictionary<int, string[]> CustomLists { get; private set; }

        public CultureInfo Culture
        {
            get; set;
        }

        public CompareOptions CompareOptions
        {
            get; set;
        }
    }
}
