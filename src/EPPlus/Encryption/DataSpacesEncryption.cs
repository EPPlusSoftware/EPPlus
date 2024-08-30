/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial licenseXml to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/29/2024         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Utils.CompundDocument;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Encryption
{
    internal class DataSpacesEncryption
    {
        internal static MemoryStream HandleDataSpaceEnryption(CompoundDocument doc)
        {
            var ms=new MemoryStream();
            var dsStorage = doc.Storage.SubStorage["\u0006DataSpaces"];
            var version = ReadVersionStream(dsStorage);
            var dataSpacemap = ReadDataSpaceMap(dsStorage);
            var dataSpaceInfo = ReadDataSpaceInfo(dsStorage.SubStorage["DataSpaceInfo"]);
            var tsInfo = dsStorage.SubStorage["TransformInfo"];
            ReadTransformReferences(tsInfo.SubStorage["DRMEncryptedTransform"]);
            var labelXml = ReadLabelXml(tsInfo);

            //Root Streams
            ReadEncryptedPropertyStreamInfo(doc);

            return ms;
        }

        private static void ReadEncryptedPropertyStreamInfo(CompoundDocument doc)
        {
            var ds = doc.Storage.DataStreams["\u0005SummaryInformation"];
            
        }

        /*
TransformInfoHeader (variable)
...
ExtensibilityHeader
XrMLLicense (variable)         
         */
        private static void ReadTransformReferences(CompoundDocument.StoragePart dsStorage)
        {
            var streamBytes = dsStorage.DataStreams["\u0006Primary"];
            using (var ms = new MemoryStream(streamBytes))
            {
                var br = new BinaryReader(ms);
                var transformLength = br.ReadInt32(); //Always 0x08
                var transformType = br.ReadInt32();
                var transformId = GetLPP4String(br, Encoding.Unicode);
                var transformName = GetLPP4String(br, Encoding.Unicode);
                var readerVersion = ReadVersion(br);
                var updateVersion = ReadVersion(br);
                var writerVersion = ReadVersion(br);

                //ExtensibilityHeader
                var el = br.ReadInt32();

                var licenseXml = GetLPP4String(br, Encoding.UTF8);
                File.WriteAllText("c:\\temp\\t.xml", licenseXml);
            }
        }

        /*
        <xsd:schema elementFormDefault = "qualified"
    xmlns:clbl="http://schemas.microsoft.com/office/2020/mipLabelMetadata"
    xmlns:r="http://schemas.microsoft.com/office/2020/02/relationships"
    xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <xsd:simpleType name = "ST_ClassificationGuid">
        < xsd:restriction base="xsd:token">
            <xsd:pattern value = "\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}" />
        </ xsd:restriction>
    </xsd:simpleType>
    <xsd:complexType name = "CT_ClassificationExtension">
        < xsd:sequence><xsd:any/></xsd:sequence><xsd:attribute name = "uri" type="xsd:token" use="required"/>
    </xsd:complexType>
    <xsd:complexType name = "CT_ClassificationExtensionList">
        < xsd:sequence><xsd:element name = "ext" type="CT_ClassificationExtension" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
    </xsd:complexType>
    <xsd:complexType name = "CT_ClassificationLabel">
        <xsd:attribute name="id" type="xsd:string" use="required"/>
        <xsd:attribute name="enabled" type="xsd:boolean" use="required"/>
        <xsd:attribute name="method" type="xsd:string" use="required"/>
        <xsd:attribute name="siteId" type="ST_ClassificationGuid" use="required"/>
        <xsd:attribute name="contentBits" type="xsd:unsignedInt" use="optional"/>
        <xsd:attribute name="removed" type="xsd:boolean" use="required"/>
    </xsd:complexType>
    <xsd:complexType name="CT_ClassificationLabelList">
        <xsd:sequence>
            <xsd:element name="label" type="CT_ClassificationLabel" minOccurs="0" maxOccurs="unbounded" />
            <xsd:element name="extLst" type="CT_ClassificationExtensionList" minOccurs="0" maxOccurs="1"/>
        </xsd:sequence>
    </xsd:complexType>
    <xsd:element name="labelList" type="CT_ClassificationLabelList" />
    </xsd:schema>
        */
        private static string ReadLabelXml(CompoundDocument.StoragePart dsStorage)
        {
            var streamBytes = dsStorage.DataStreams["LabelInfo"];
            using (var ms = new MemoryStream(streamBytes))
            {
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }
        private static object ReadDataSpaceInfo(CompoundDocument.StoragePart dsStorage)
        {
            var streamBytes = dsStorage.DataStreams["DRMEncryptedDataSpace"];
            using (var ms = new MemoryStream(streamBytes))
            {
                var br = new BinaryReader(ms);
                var headerLength = br.ReadInt32(); //Always 0x08
                var entryCount = br.ReadInt32();
            }
            return null;
        }
        private static string ReadVersionStream(CompoundDocument.StoragePart dsStorage)
        {
            var streamBytes = dsStorage.DataStreams["Version"];
            using (var ms = new MemoryStream(streamBytes))
            {
                var br = new BinaryReader(ms);
                var name = GetLPP4String(br, Encoding.Unicode);
                var readerVersion = ReadVersion(br);
                var updaterVersion = ReadVersion(br);
                var writerVersion = ReadVersion(br);
                return name + "," + readerVersion + "," + updaterVersion + "," + writerVersion;
            }
        }

        private static string ReadVersion(BinaryReader br)
        {
            var major = br.ReadUInt16(); 
            var minor = br.ReadUInt16();

            return $"{major}.{minor}";
        }

        private static string GetLPP4String(BinaryReader br, Encoding enc)
        {
            var length = br.ReadInt32();
            var data = new byte[length];
            br.Read(data, 0, length);
            if (length % 4 != 0)
            {
                br.ReadBytes(length % 4);
            }
            //Padding??vf
            return enc.GetString(data, 0, length);
        }

        private static List<DataSpaceReference> ReadDataSpaceMap(CompoundDocument.StoragePart dsStorage)
        {
            var l=new List<DataSpaceReference>();
            var streamBytes = dsStorage.DataStreams["DataSpaceMap"];
            using (var ms = new MemoryStream(streamBytes))
            {
                var br = new BinaryReader(ms);
                var headerLength = br.ReadInt32(); //Always 0x08
                var entryCount = br.ReadInt32();
                for (int i = 0; i < entryCount; i++) 
                {
                    var length = br.ReadInt32();
                    var referenceComponentCount=br.ReadInt32();
                    for(int j = 0;j < referenceComponentCount;j++)
                    {
                        var r=new DataSpaceReference();
                        r.ReferenceType = (DataSpaceReference.eReferenceType)br.ReadInt32();
                        r.Name = GetLPP4String(br, Encoding.UTF8);
                        l.Add(r);
                    }
                }
            }
            return l;
        }
    }

    internal class DataSpaceReference
    {
        internal enum eReferenceType
        {
            Stream=0,
            Storage=1
        }
        public eReferenceType ReferenceType { get; set; }
        public string Name { get; set; }
    }
}
