using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Metadata;
using OfficeOpenXml.Metadata.FutureMetadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static OfficeOpenXml.ExcelWorksheet;

namespace OfficeOpenXml.RichData
{
    internal class RichDataStore
    {
        public RichDataStore(ExcelWorksheet sheet)
        {
            _sheet = sheet;
            _workbook = sheet.Workbook;
            _metadataStore = sheet._metadataStore;
            _metadata = sheet.Workbook.Metadata;
        }

        private readonly ExcelWorksheet _sheet;
        private readonly ExcelWorkbook _workbook;
        private readonly CellStore<MetaDataReference> _metadataStore;
        private readonly ExcelMetadata _metadata;

        internal ExcelRichValue GetRichData(int row, int col, string structureType = null)
        {
            var vm = _metadataStore.GetValue(row, col).vm;
            if (vm == 0 || !_metadata.IsRichData(vm)) return null;
            // vm is a 1-based index pointer
            var vmIx = vm - 1;
            var valueMd = _metadata.ValueMetadata[vmIx];
            var valueRecord = valueMd.Records.First();
            var type = _metadata.MetadataTypes[valueRecord.RecordTypeIndex - 1];
            var futureMetadata = _metadata.MetadataTypes.First(x => x.Name == type.Name);
            var rdv = _workbook.RichData.Values.Items[valueRecord.ValueTypeIndex];
            if(!string.IsNullOrEmpty(structureType) && structureType != rdv.Structure.Type)
            {
                return null;
            }
            return rdv;
        }

        internal void AddRichData(string relationshipType, string target, IEnumerable<string> values, RichDataStructureFlags structureFlag, out int vmIndex)
        {
            _workbook.RichData.RichValueRels.AddItem(target, relationshipType, out int relIx);
            var structureId = _workbook.RichData.Structures.GetStructureId(structureFlag);
            var rv = new ExcelRichValue(structureId)
            {
                Structure = _workbook.RichData.Structures.StructureItems[structureId]
            };
            if((structureFlag & RichDataStructureFlags.LocalImage) == RichDataStructureFlags.LocalImage)
            {
                rv.AddLocalImage(relIx, int.Parse(values.ElementAt(0)), string.Empty);
            }
            _workbook.RichData.Values.Items.Add(rv);

            // update the metadata
            _metadata.CreateRichValueMetadata(_workbook.RichData, out int vm);
            vmIndex = vm;
        }

        internal RichValueRel GetRelation(int relationIndex)
        {
            return _workbook.RichData.GetRelation(relationIndex);
        }

        internal RichValueRel GetRelation(string target, string type)
        {
            return _workbook.RichData.GetRelation(target, type);
        }

    }
}
