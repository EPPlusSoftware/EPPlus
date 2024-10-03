using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData
{
    internal class ExcelRichDataDeletions
    {
        public ExcelRichDataDeletions()
        {
            _metadataIndexes = new HashSet<int>();
            _richDataIndexes = new HashSet<int>();
        }

        private readonly HashSet<int> _metadataIndexes;
        private readonly HashSet<int> _richDataIndexes;

        internal bool IsEmpty
        {
            get
            {
                return _metadataIndexes.Count == 0 && _richDataIndexes.Count == 0;
            }
        }

        internal bool IsDeleted(int metadataIndex, int richDataIndex)
        {
            return _metadataIndexes.Contains(metadataIndex) && _richDataIndexes.Contains(richDataIndex);
        }

        internal bool DeleteRichData(int metadataIndex, int richDataIndex)
        {
            if(_metadataIndexes.Contains(metadataIndex) || _richDataIndexes.Contains(richDataIndex))
            {
                return false;
            }
            _metadataIndexes.Add(metadataIndex); 
            _richDataIndexes.Add(richDataIndex);
            return true;
        }

        internal bool RemoveDeletion(int metadataIndex, int richDataIndex)
        {
            if (!(_metadataIndexes.Contains(metadataIndex) && _richDataIndexes.Contains(richDataIndex)))
            {
                return false;
            }
            _metadataIndexes.Remove(metadataIndex);
            _richDataIndexes.Remove(richDataIndex);
            return true;
        }

        public IEnumerable<int> GetSortedMetadataIndexes()
        {
            return _metadataIndexes.ToList().OrderBy(x => x);
        }

        public IEnumerable<int> GetSortedRichdataIndexes()
        {
            return _richDataIndexes.ToList().OrderBy(x => x);
        }
    }
}
