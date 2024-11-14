/*************************************************************************************************
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
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.RichData.RichValueArrays
{
    internal class ExcelRichDataArrayCollection : IndexedCollection<ExcelRichDataArray>
    {
        const string PART_URI_PATH = "/xl/richData/rdarray.xml";
        private readonly Uri _uri;
        private ExcelWorkbook _wb;
        private readonly ExcelRichData _richData;
        private readonly RichDataIndexStore _indexStore;
        ZipPackagePart _part;
        internal ZipPackagePart Part { get { return _part; } }

        public ExcelRichDataArrayCollection(ExcelWorkbook wb, ExcelRichData richData) : base(wb.IndexStore, RichDataEntities.RichDataArray)
        {
            _wb = wb;
            _richData = richData;
            _indexStore = wb.IndexStore;
            var r = wb.Part.GetRelationshipsByType(Relationsships.schemaRichDataRichDataValueArray).FirstOrDefault();
            if (r == null)
            {
                _uri = new Uri(PART_URI_PATH, UriKind.Relative);
            }
            else
            {
                _uri = UriHelper.ResolvePartUri(r.SourceUri, r.TargetUri);
            }
            LoadPart(wb);
        }

        private void LoadPart(ExcelWorkbook wb)
        {
            if (wb._package.ZipPackage.PartExists(_uri))
            {
                _part = wb._package.ZipPackage.GetPart(_uri);
                ReadXml(_part.GetStream());
            }
        }

        private void ReadXml(Stream stream)
        {
            var xr = XmlReader.Create(stream);
            while (xr.Read())
            {
                if (xr.IsElementWithName("a"))
                {
                    var array = new ExcelRichDataArray(_richData, _indexStore, xr);
                    var rvId = _richData.Values.GetIdByIndex((int)array.RichValueId);
                    var rv = _richData.Values.Get(rvId);
                    rv.AddRelationTo(array);
                    array.RichValueId = rv.Id;
                    Add(array);
                }
            }

        }
    }
}
