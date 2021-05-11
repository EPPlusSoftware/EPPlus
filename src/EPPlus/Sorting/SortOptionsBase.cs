using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sorting
{
    public abstract class SortOptionsBase
    {
        public SortOptionsBase()
        {
            ColumnIndexes = new List<int>();
            Descending = new List<bool>();
            CustomLists = new Dictionary<int, string[]>();
            CompareOptions = CompareOptions.None;
        }

        internal List<int> ColumnIndexes { get; private set; }
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
