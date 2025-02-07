/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/7/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.Worksheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sorting.Internal
{
    internal class RangeWorksheetData
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public RangeWorksheetData(ExcelRangeBase range)
        {
            var worksheet = range.Worksheet;
            
            //Shared formulas spaning multiple cells will be converted to single cell formulas.
            ConvertSharedFormulas(range);

            Flags = GetItems(range, worksheet._flags);
            Formulas = GetItems(range, worksheet._formulas);
            Hyperlinks = GetItems(range, worksheet._hyperLinks);
            Comments = GetItems(range, worksheet._commentsStore);
            Metadata = GetItems(range, worksheet._metadataStore);
        }

        private void ConvertSharedFormulas(ExcelRangeBase range)
        {
            var cse = new CellStoreEnumerator<object>(range._worksheet._formulas, range);
            var hs = new HashSet<int>();
            foreach (var item in cse)
            {
                if(cse.Value is int sfIx && hs.Contains(sfIx)==false)
                {
                    hs.Add(sfIx);
                }
            }
            foreach (var x in hs)
            {
                var sf = range._worksheet._sharedFormulas[x];
                if(sf.IsSingleCell==false)
                {
                    if(sf.IsArray)
                    {
                        throw new InvalidOperationException($"Can't sort part of an array: {sf.Address}");
                    }
                    WorksheetRangeHelper.ConvertSharedFormulaToCellFormula(range._worksheet,sf, new ExcelAddressBase(sf.Address));
                    range._worksheet._sharedFormulas.Remove(x);
                }
            }
        }

        public Dictionary<string, byte> Flags { get; private set; }

        public Dictionary<string, object> Formulas { get; private set; }

        public Dictionary<string, Uri> Hyperlinks { get; private set; }

        public Dictionary<string, int> Comments { get; private set; }

        public Dictionary<string, ExcelWorksheet.MetaDataReference> Metadata { get; private set; }

        private static Dictionary<string, T> GetItems<T>(ExcelRangeBase r, CellStore<T> store)
        {
            var e = new CellStoreEnumerator<T>(store, r._fromRow, r._fromCol, r._toRow, r._toCol);
            var l = new Dictionary<string, T>();
            while (e.Next())
            {
                l.Add(e.CellAddress, e.Value);
            }
            return l;
        }
    }
}
