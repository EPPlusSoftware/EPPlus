﻿using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using OfficeOpenXml.Utils.Extensions;
using System.Runtime.InteropServices;
using OfficeOpenXml.Packaging.Ionic.Zip;

namespace OfficeOpenXml.RichData.Types
{
    internal class ExcelRichDataValueTypeInfo
    {
        private ExcelWorkbook _wb;
        private Uri _uri=null;
        private ZipPackagePart _part=null;
        private const string PART_URI_PATH = "/xl/richData/rdRichValueTypes.xml";
        public ExcelRichDataValueTypeInfo(ExcelWorkbook wb)
        {
            _wb = wb;
            _uri = new Uri(PART_URI_PATH, UriKind.Relative);
            ReadPart(wb);
        }
        public ExcelRichDataValueTypeInfo(ExcelWorkbook wb, ZipPackageRelationship r) 
        {
            _wb = wb;
            if (r != null)
            {
                _uri = UriHelper.ResolvePartUri(r.SourceUri, r.TargetUri);
                ReadPart(wb);
            }
        }

        private void ReadPart(ExcelWorkbook wb)
        {
            if (wb._package.ZipPackage.PartExists(_uri))
            {
                _part = wb._package.ZipPackage.GetPart(_uri);
                ReadXml(_part.GetStream());
            }
        }

        internal ZipPackagePart Part { get { return _part; } }
        private void ReadXml(Stream stream)
        {
            var xr = XmlReader.Create(stream);
            while(xr.Read())
            {
                if(xr.IsElementWithName("global"))
                {
                    ReadKeyFlags(xr, Global, out _globalExtLstXml);
                }
                else if(xr.IsElementWithName("types"))
                {
                    ReadKeyFlags(xr, Types, out _typesExtLstXml);
                }
                else if(xr.IsElementWithName("extLst"))
                {
                    ExtLstXml = xr.ReadInnerXml();
                }
                else if(xr.IsEndElementWithName("rvTypesInfo"))
                {
                    break;
                }
            }
        }
        private void ReadKeyFlags(XmlReader xr, Dictionary<string, ExcelRichTypeValueKey> values, out string extLst)
        {
            if(xr.ReadUntil("key", "extLst"))
            {
                ReadValues(xr, values);
                if (xr.IsElementWithName("extLst"))
                {
                    extLst= xr.ReadInnerXml();
                }
                else
                {
                    extLst = null;
                }
                return;
            }
            else
            {
                extLst = null;
            }
        }

        private void ReadValues(XmlReader xr, Dictionary<string, ExcelRichTypeValueKey> values)
        {
            while(xr.IsElementWithName("key") && xr.EOF == false)
            {
                while(!xr.IsEndElementWithName("key") && xr.EOF==false)
                {
                    var item = new ExcelRichTypeValueKey(xr.GetAttribute("name"));
                    values.Add(item.Name, item);
                    while (xr.Read())
                    {
                        if(xr.IsElementWithName("flag"))
                        {
                            var flag = xr.GetAttribute("name").ToEnum<RichValueKeyFlags>();
                            if (flag.HasValue)
                            {
                                var v = xr.GetAttribute("value");
                                if (v == "1" || v.Equals("true", StringComparison.OrdinalIgnoreCase))
                                {
                                    item.Flags |= flag.Value;
                                }
                                else
                                {
                                    item.Flags &= ~flag.Value;
                                }
                            }
                        }
                        else
                        {
                            if (xr.NodeType == XmlNodeType.EndElement) xr.Read();
                            if (xr.Name != "flag") break;
                        }
                    }
                    if(xr.IsEndElementWithName("keyFlags"))
                    {
                        xr.Read(); //Move to global/types end element
                        return;
                    }
                }
            }
        }
        internal void CreatePart()
        {
            if (Global.Count == 0 && Types.Count == 0 && ExtLstXml == null) return;
            if (_part == null)
            {
                _uri = new Uri(PART_URI_PATH, UriKind.Relative);
                _part = _wb._package.ZipPackage.CreatePart(_uri, ContentTypes.contentTypeRichDataValueType);
                _wb.Part.CreateRelationship(_uri, TargetMode.Internal, Relationsships.schemaRichDataValueTypeRelationship);
            }
            _part.SaveHandler = Save;
        }
        internal void Save(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
        {
            stream.PutNextEntry(fileName);
            stream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
            var sw = new StreamWriter(stream);
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write($"<rvTypesInfo xmlns=\"{Schemas.schemaRichData2}\" xmlns:mc=\"{Schemas.schemaMarkupCompatibility}\" xmlns:x=\"{ExcelPackage.schemaMain}\" mc:Ignorable=\"x\">");
            WriteSection(sw, Global, "global", _globalExtLstXml);
            WriteSection(sw, Types, "types", _typesExtLstXml);
            WriteExtLst(sw, ExtLstXml);
            sw.Write("</rvTypesInfo>");
            sw.Flush();
        }

        private void WriteExtLst(StreamWriter sw, string extLstXml)
        {
            if (string.IsNullOrEmpty(ExtLstXml)==false)
            {
                sw.Write("<extLst>");
                sw.Write(extLstXml);
                sw.Write("</extLst>");
            }
        }

        private void WriteSection(StreamWriter sw, Dictionary<string, ExcelRichTypeValueKey> section, string elementName, string extLstXml)
        {
            if (section.Count > 0)
            {
                sw.Write($"<{elementName}><keyFlags>");
                foreach (var item in Global.Values)
                {
                    item.WriteXml(sw);
                }
                sw.Write($"</keyFlags></{elementName}>");
                WriteExtLst(sw, extLstXml);
            }
        }

        internal void CreateDefault()
        {            
            Global.Add("_self", new ExcelRichTypeValueKey("_Self") { Flags = RichValueKeyFlags.ExcludeFromFile | RichValueKeyFlags.ExcludeFromCalcComparison });
            Global.Add("_DisplayString", new ExcelRichTypeValueKey("_DisplayString") { Flags = RichValueKeyFlags.ExcludeFromCalcComparison });
            Global.Add("_Flags", new ExcelRichTypeValueKey("_Flags") { Flags = RichValueKeyFlags.ExcludeFromCalcComparison });
            Global.Add("_Format", new ExcelRichTypeValueKey("_Format") { Flags = RichValueKeyFlags.ExcludeFromCalcComparison });
            Global.Add("_SubLabel", new ExcelRichTypeValueKey("_SubLabel") { Flags = RichValueKeyFlags.ExcludeFromCalcComparison });
            Global.Add("_Attribution", new ExcelRichTypeValueKey("_Attribution") { Flags = RichValueKeyFlags.ExcludeFromCalcComparison });
            Global.Add("_Icon", new ExcelRichTypeValueKey("_Icon") { Flags = RichValueKeyFlags.ExcludeFromCalcComparison });
            Global.Add("_Display", new ExcelRichTypeValueKey("_Display") { Flags = RichValueKeyFlags.ExcludeFromCalcComparison });
            Global.Add("_CanonicalPropertyNames", new ExcelRichTypeValueKey("_CanonicalPropertyNames") { Flags = RichValueKeyFlags.ExcludeFromCalcComparison });
            Global.Add("_ClassificationId", new ExcelRichTypeValueKey("_ClassificationId") { Flags = RichValueKeyFlags.ExcludeFromCalcComparison });
        }
        public Dictionary<string, ExcelRichTypeValueKey>  Global { get; set; } = new Dictionary<string, ExcelRichTypeValueKey>();
        public Dictionary<string, ExcelRichTypeValueKey> Types { get; set; } = new Dictionary<string, ExcelRichTypeValueKey>();
        public string ExtLstXml { get; set; }
        private string _globalExtLstXml=null, _typesExtLstXml=null;        
    }
}
