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
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Utils.CompundDocument;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Xml;
using System.Linq;
using System.Threading;
using System.Security.Cryptography;
using OfficeOpenXml.Utils;
using System.Security.Cryptography.X509Certificates;
using OfficeOpenXml.Interfaces;
using OfficeOpenXml.SensitivityLabels;
namespace OfficeOpenXml.Encryption
{
    internal class DataSpacesEncryption
    {
        internal static SensibilityLabelInfo ReadDataSpaceEnrcyptionInfo(CompoundDocument doc)
        {
            try
            {
                var si = new SensibilityLabelInfo();
                var dsStorage = doc.Storage.SubStorage["\u0006DataSpaces"];
                si.Version = ReadVersionStream(dsStorage);
                si.DataSpaceMap = ReadDataSpaceMap(dsStorage);
                si.DataSpaceInfo = ReadDataSpaceInfo(dsStorage.SubStorage["DataSpaceInfo"]);
                var tsInfo = dsStorage.SubStorage["TransformInfo"];
                si.Transformation = ReadTransformReferences(tsInfo.SubStorage["DRMEncryptedTransform"]);
                si.LabelXml = ReadLabelXml(tsInfo);

                //Root Streams
                si.SummaryInfoProperties = ReadEncryptedPropertyStreamInfo(doc.Storage.DataStreams["\u0005SummaryInformation"]);

                si.SummaryDocumentInfoProperties = ReadEncryptedPropertyStreamInfo(doc.Storage.DataStreams["\u0005DocumentSummaryInformation"]);

                return si;
            }
            catch (Exception ex)
            {
                throw(new InvalidDataException("EPPlus can not read this package.", ex));
            }
        }

        private static MemoryStream DecryptPackage(CompoundDocument doc, TransformInfoHeader transformInfo)
        {
            var ms = new MemoryStream();
            var stream = doc.Storage.DataStreams["EncryptedPackage"];
            var br = new BinaryReader(new MemoryStream(stream));
            var size = br.ReadUInt64();
            var er = GetEI(transformInfo);
            ms.Write(br.ReadBytes((int)size), 0, (int)size);
            ms.Flush();
            return ms;
        }

        private static EncryptionInfo GetEI(TransformInfoHeader transformInfo)
        {
            var settings = new XmlReaderSettings { ConformanceLevel = ConformanceLevel.Fragment }; //This XrML has multiple root elements.
#if (NET35)
            settings.ProhibitDtd = true;
#else
            settings.DtdProcessing = DtdProcessing.Prohibit;
#endif
            var xr = XmlReader.Create(new StringReader(transformInfo.LicenseXrML), settings);
            while (xr.Read())
            {
                if(xr.LocalName== "XrML" && xr.NodeType==XmlNodeType.Element && xr.IsEmptyElement==false)
                {                    
                    var data = ReadXrMLLicense(xr);
                }
            }
            return new EncryptionInfoBinary();
        }

        private static object ReadXrMLLicense(XmlReader xr)
        {
            while (xr.Read())
            {                 
                if(xr.LocalName == "XrML" && xr.NodeType==XmlNodeType.Element)
                {
                    break;
                }
                switch (xr.LocalName)
                {
                    case "PUBLICKEY":
                        //var pr = GetPublicKey(xr);
                        break;
                    case "AUTHENTICATEDDATA":
                        var decryptedData = GetDecryptAuthData(xr);
                        break;
                }
            }
            return null;
        }

        private static object GetDecryptAuthData(XmlReader xr)
        {
            if(xr.GetAttribute("id") == "Encrypted-Rights-Data")
            {
                var key = Convert.FromBase64String("cffQtZqwEfY+PtHy8jH14FPGDz2phwzbqGYqn/GDaezTAmBHL4N61AAHmKKPY7tejtU/a7RiuNvPs5GayUjhfpGyJBUrX23rfWImnenCfa1oaqZAnfxZ/DoML4jTdrlF+59XTLsVmgkVB66jpquz9KX9WmDAFqqCc00N5TcBBanu9i+gotBNlyZ/pxSV/+tEMZSzefa+kJauYF/2kidNV4yLGD7PCMH4E1LqOkM0t7+jgdC73ymFFFRI14GT2pc/G4sYKawnrX22lX3s/9Gx1EHnJZRluitLObUpwzAoN12D9th44zgGjRMmzJP2K4Ov2mG+gVtF/qs2l93H713+nw==");
                var content = xr.ReadElementContentAsString();
                var data = Convert.FromBase64String(content);

                //using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider(2048))
                //{
                //    rsa.ImportRSAPrivateKey(key, out _);
                //    return rsa.Decrypt(data, false); // Use false for PKCS#1 v1.5 padding
                //}

                //var s= DecryptAES(data, key);
                using (var rsaDecryptor = new RSACryptoServiceProvider(2048))
                {
                    //rsaDecryptor.PersistKeyInCsp = false;

                    // Import the private key (use the same private key that was used for encryption)

                    var keyInfo = new RSAParameters
                    {
                        Modulus = key,
                        Exponent = [(byte)0x1, (byte)0x0, (byte)0x1]
                    };
                    rsaDecryptor.ImportParameters(keyInfo);

                    // Decrypt the data
                    var decryptedData = rsaDecryptor.Decrypt(data, false); // Use false for PKCS#1 v1.5 padding
                }
                //var rsa = new RSACryptoServiceProvider(keyInfo);
                //rsa.KeySize = 256;

                //rsa.Decrypt(data)
            }
            return null;
        }
        static string DecryptAES(byte[] data, byte[] key)
        {
            using (Aes aes = Aes.Create())
            {
                aes.KeySize = 2048;
                aes.Key = key;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                // Assume the IV is prepended to the data. Otherwise, you would need to know how the IV is handled.
                byte[] iv = new byte[aes.BlockSize / 8];
                byte[] cipherText = new byte[data.Length - iv.Length];

                Array.Copy(data, iv, iv.Length);
                Array.Copy(data, iv.Length, cipherText, 0, cipherText.Length);

                aes.IV = iv;

                using (var decryptor = aes.CreateDecryptor(aes.Key, aes.IV))
                {
                    byte[] decryptedData = decryptor.TransformFinalBlock(cipherText, 0, cipherText.Length);
                    return Encoding.UTF8.GetString(decryptedData);
                }
            }
        }
        private static List<object> ReadEncryptedPropertyStreamInfo(byte[] streamBytes)
        {
            using (var ms = new MemoryStream(streamBytes))
            {
                var br = new BinaryReader(ms);
                //See MS-OLEPS section 2.21 & 2.23
                var byteOrder = br.ReadUInt16(); //Always 0xFFFE
                var version = br.ReadUInt16(); //Property version - 0 or 1
                var systemIdentifier = br.ReadUInt32();
                var clsid = new Guid(br.ReadBytes(16));
                var numPropertySets = br.ReadInt32();
                var fMTID0GUID = new Guid(br.ReadBytes(16));
                var offset0 = br.ReadUInt32();
                if (numPropertySets == 2)
                {
                    var fMTID1GUID = new Guid(br.ReadBytes(16));
                    var offset1 = br.ReadUInt32();
                }
                ms.Position = offset0;
                var ps1 = ReadPropertySet(br);
                return ps1;
                //if(numPropertySets==2)
                //{
                //    var ps2 = ReadPropertySet(br);
                //}
            }

        }
        private class PropertyOffset
        {
            public PropertyOffset(int identifyer, int offset)
            {                    
                Identifier = identifyer;
                Offset = offset;
            }
            public int Identifier { get; set; }
            public int Offset { get; set; }
        }
        private static List<object> ReadPropertySet(BinaryReader br)
        {
            var startPos = br.BaseStream.Position;
            var size = br.ReadInt32();
            var numberOfProperties = br.ReadInt32();
            var offsets = new List<PropertyOffset>();
            for(var i = 0;i < numberOfProperties;i++)
            {
                var propertyIdentifier = br.ReadInt32();
                var offset = br.ReadInt32();
                offsets.Add(new PropertyOffset(propertyIdentifier, offset));

            }

            var properties = new List<object>();
            Encoding encoding = null;
            foreach(var propertyOffset in offsets)
            {
                if(propertyOffset.Identifier==1)
                {
                    encoding = GetEncoding(br);
                }
                if (propertyOffset.Identifier > 1 && propertyOffset.Identifier < 0x7FFFFFFF)
                {
                    properties.Add(GetProperty(propertyOffset, br, encoding));
                }
                else
                {
                    //Directory, Codepage, Local or Behavior
                }
            }

            return properties;
        }

        private static Encoding GetEncoding(BinaryReader br)
        {
            var id = br.ReadInt32(); // Should be 2
            var cp = br.ReadInt32();
            return Encoding.GetEncoding(cp);
        }

        private static object GetProperty(PropertyOffset propertyOffset, BinaryReader br, Encoding encoding)
        {
            var type = br.ReadUInt16();
            br.ReadBytes(2); //Padding
            object value = null;
            switch (type) 
            {
                case 0x2:
                    value = br.ReadInt16();
                    br.ReadBytes(2); //Padding
                    break;
                case 0x3: 
                    value = br.ReadInt32();
                    break;
                case 0x4:
                    value = BitConverter.ToSingle(br.ReadBytes(4),0);
                    break;
                case 0x5:
                    value = BitConverter.ToDouble(br.ReadBytes(8), 0);
                    break;
                case 0xb:
                    value = br.ReadInt32() != 0;
                    break;
                case 0x1e:
                    value = ReadString(br, encoding);
                    break;
                case 0x40:
                    var lowDT = br.ReadUInt32();
                    var highDT = br.ReadUInt32();
                    long ns = (long)lowDT | (long)highDT<<32;
                    value = new DateTime(1601, 1, 1).AddTicks(ns);
                    break;
                case 0x100C:
                    var count = br.ReadUInt32();
                    count /= 2;
                    var array = new string[count];
                    for (int i = 0; i < count; i++)
                    {
                        var st=br.ReadUInt16();
                        br.ReadUInt16();
                        if (st==0x1E)
                        {
                            array[i] = ReadString(br, encoding);
                        }
                        else
                        {
                            array[i] = ReadString(br, Encoding.Unicode);
                        }
                    }
                    value = array;
                    break;
                case 0x101E:
                    var elements = br.ReadUInt32();
                    array = new string[elements];
                    for(int i= 0; i < elements;i++)
                    {
                        array[i]=ReadString(br, encoding);
                    }
                    value = array;
                    break;
            }
            return value;
        }

        private static string ReadString(BinaryReader br, Encoding encoding)
        {
            var size = br.ReadInt32();
            var bytes = br.ReadBytes(size);
            size = Array.FindIndex(bytes, x => x == 0);
            return encoding.GetString(bytes, 0, size);
        }

        /*
TransformInfoHeader (variable)
...
ExtensibilityHeader
XrMLLicense (variable)         
         */
        internal class TransformInfoHeader
        {
            public int TransformType { get; set; }
            public string TransformId { get; set; }
            public string TransformName { get; set; }
            public string ReaderVersion { get; set; }
            public string UpdaterVersion { get; set; }
            public string WriterVersion { get; set; }
            public string LicenseXrML { get; set; }
        }
        private static TransformInfoHeader ReadTransformReferences(CompoundDocument.StoragePart dsStorage)
        {
            var streamBytes = dsStorage.DataStreams["\u0006Primary"];
            var ih=new TransformInfoHeader();
            using (var ms = new MemoryStream(streamBytes))
            {
                var br = new BinaryReader(ms);
                var transformLength = br.ReadInt32(); //Always 0x08
                ih.TransformType = br.ReadInt32();
                ih.TransformId = GetLPP4String(br, Encoding.Unicode);
                ih.TransformName = GetLPP4String(br, Encoding.Unicode);
                ih.ReaderVersion = ReadVersion(br);
                ih.UpdaterVersion = ReadVersion(br);
                ih.WriterVersion = ReadVersion(br);

                //ExtensibilityHeader
                var el = br.ReadInt32(); //Currently not used. Should always be 4.

                ih.LicenseXrML = GetLPP4String(br, Encoding.UTF8);
            }
            return ih;
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
        private static List<string> ReadDataSpaceInfo(CompoundDocument.StoragePart dsStorage)
        {
            var l = new List<string>();
            var streamBytes = dsStorage.DataStreams["DRMEncryptedDataSpace"];
            using (var ms = new MemoryStream(streamBytes))
            {
                var br = new BinaryReader(ms);
                var headerLength = br.ReadInt32(); //Always 0x08
                var entryCount = br.ReadInt32();
                for (int i = 0; i < entryCount; i++)
                {
                    l.Add(GetLPP4String(br, Encoding.Unicode));
                }
            }
            return l;
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
                        r.Name = GetLPP4String(br, Encoding.Unicode);
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
