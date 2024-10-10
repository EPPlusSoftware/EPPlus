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
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.RichData.Structures.SupportingPropertyBags
{
    internal class SupportingPropertyBags
    {
        private ExcelWorkbook _wb;
        private ZipPackagePart _part;
        private Uri _uri;
        private const string PART_URI_PATH = "/xl/richData/rdsupportingpropertybag.xml";
        private List<SupportingPropertyBagArray> _arrays = new List<SupportingPropertyBagArray>();
        private List<SupportingPropertyBagData> _data = new List<SupportingPropertyBagData>();

        internal SupportingPropertyBags(ExcelWorkbook wb)
        {
            _wb = wb;
            var r = wb.Part.GetRelationshipsByType(Relationsships.schemaRichDataSupportingPropertyBagRelationship).FirstOrDefault();
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
                if (xr.IsElementWithName("spbArrays"))
                {
                    ReadArrays(xr);
                }
                else if(xr.IsElementWithName("spbData"))
                {
                    ReadData(xr);
                }
            }
        }

        private void ReadArrays(XmlReader xr)
        {
            _arrays.Clear();
            while (xr.Read())
            {
                if(xr.IsElementWithName("a"))
                {
                    var arr = SupportingPropertyBagArray.CreateFromXml(xr);
                    _arrays.Add(arr);
                }
                if (xr.IsEndElementWithName("spbArrays"))
                {
                    break;   
                }
            }
        }

        private void ReadData(XmlReader xr)
        {
            _data.Clear();
            while (xr.Read())
            {
                var data = SupportingPropertyBagData.CreateFromXml(xr);
                if(data != null)
                {
                    _data.Add(data);
                }
                if (xr.IsEndElementWithName("spbData"))
                {
                    break;
                }
            }
        }

        internal void Save(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
        {
            stream.PutNextEntry(fileName);
            stream.CompressionLevel = (Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
            var sw = new StreamWriter(stream);

            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write($"<supportingPropertyBags xmlns=\"{Schemas.schemaRichData}\">");
            if(_arrays.Count > 0)
            {
                sw.Write($"<spbArrays count=\"{_arrays.Count}\">");
                foreach (var array in _arrays)
                {
                    array.WriteXml(sw);
                }
                sw.Write("</spbArrays>");
            }
            if(_data.Count > 0)
            {
                sw.Write($"<spbData count=\"{_arrays.Count}\">");
                foreach(var data in _data)
                {
                    data.WriteXml(sw);
                }
                sw.Write("</spbData>");
            }
            sw.Write("</supportingPropertyBags>");
            sw.Flush();
        }
    }
}
