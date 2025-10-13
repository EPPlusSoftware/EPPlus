/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
using OfficeOpenXml.Packaging;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// Settings for power query connections.
    /// </summary>
    public class ExcelPowerQuerySettings
    {
        internal ExcelPowerQuerySettings()
        {
            PermissionsXml = new XmlDocument();
            XmlHelper.LoadXmlSafe(PermissionsXml, "<?xml version=\"1.0\" encoding=\"utf-8\"?><PermissionList xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><CanEvaluateFuturePackages>false</CanEvaluateFuturePackages><FirewallEnabled>true</FirewallEnabled></PermissionList>", Encoding.UTF8);
            MetadataXml = new XmlDocument();
            XmlHelper.LoadXmlSafe(MetadataXml, "<?xml version=\"1.0\" encoding=\"utf-8\"?><LocalPackageMetadataFile xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><Items><Item><ItemLocation><ItemType>AllFormulas</ItemType><ItemPath /></ItemLocation><StableEntries><Entry Type=\"Relationships\" Value=\"sAAAAAA==\" /></StableEntries></Item></Items></LocalPackageMetadataFile>", Encoding.UTF8);
            PowerQueryFormulas = "section Section1;\n";
        }
        internal ExcelPowerQuerySettings(byte[] blob)
        {
            using var ms = new MemoryStream(blob);
            var br = new BinaryReader(ms);
            var version = br.ReadInt32();
            var size = br.ReadInt32();
            var pck = br.ReadBytes((int)size);
            PQPackage = new ZipPackage(new MemoryStream(pck));

            var section1MPart = PQPackage.GetPartByContentType("application/x-ms-m");
            using var mms = section1MPart.GetStream();
            using (var reader = new StreamReader(mms, Encoding.UTF8))
            {
                PowerQueryFormulas = reader.ReadToEnd();
            }

            size = br.ReadInt32();
            var permBytes = br.ReadBytes((int)size);
            var permissionXml = Encoding.UTF8.GetString(permBytes);
            PermissionsXml = new XmlDocument();
            XmlHelper.LoadXmlSafe(PermissionsXml, permissionXml, Encoding.UTF8);

            size = br.ReadInt32();
            version = br.ReadInt32();
            size = br.ReadInt32();
            var metadataXml = Encoding.UTF8.GetString(br.ReadBytes((int)size));
            size = br.ReadInt32();
            MetadataXml = new XmlDocument();
            XmlHelper.LoadXmlSafe(MetadataXml, metadataXml, Encoding.UTF8);

            MetadataContentPackage = br.ReadBytes((int)size);
            size = br.ReadInt32();

            var packageBinding = br.ReadBytes((int)size);
#if(!Core)            
            var protectedData = ProtectedData.Unprotect(packageBinding, UTF8Encoding.UTF8.GetBytes("DataExplorer Package Components"), DataProtectionScope.CurrentUser);
            br = new BinaryReader(new MemoryStream(protectedData));
            size = br.ReadInt32();
            var hash1 = br.ReadBytes(size);
            size = br.ReadInt32();
            var hash2 = br.ReadBytes(size);

            var sha = SHA256.Create();
            var calcHash1 = sha.ComputeHash(pck);
            var calcHash2 = sha.ComputeHash(permBytes);
#endif
        }
        internal void Save()
        {
            ZipPackagePart sectionMPart;
            if(PQPackage==null)
            {
                PQPackage=new ZipPackage(new MemoryStream());
                var pp = PQPackage.CreatePart(new Uri("/Config/Package.xml", UriKind.Relative), "text/xml");
                var sw = new StreamWriter(pp.GetStream(), Encoding.UTF8); 
                sw.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?><Package xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><Version>2.147.503.0</Version><MinVersion>2.21.0.0</MinVersion><Culture>en-SE</Culture></Package>");
                sw.Flush();
                sectionMPart = PQPackage.CreatePart(new Uri("/Formulas/Section1.m", UriKind.Relative), "application/x-ms-m");
            }
            else
            {
                sectionMPart = PQPackage.GetPartByContentType("application/x-ms-m");
                if(sectionMPart==null)
                {
                    sectionMPart = PQPackage.CreatePart(new Uri("/Formulas/Section1.m", UriKind.Relative), "application/x-ms-m");
                }
            }
            using var ms = sectionMPart.GetStream();
            var streamwriter = new StreamWriter(ms, Encoding.UTF8);
            streamwriter.Write(PowerQueryFormulas);
            streamwriter.Flush();
            streamwriter.Close();

            var packageStream = new MemoryStream();
            PQPackage.Save(packageStream);
            var pckbytes = packageStream.ToArray();

            var bw= new BinaryWriter(new MemoryStream());
            bw.Write(0);
            bw.Write(pckbytes.Length);
            bw.Write(pckbytes);
            var permBytes = Encoding.UTF8.GetBytes(PermissionsXml.OuterXml);

            var metadataBytes = GetMetaDataBytes();
            bw.Write(metadataBytes.Length);
            bw.Write(metadataBytes);

            // Permission binding. We set it to empty as DPAPI is only available on Windows.
            bw.Write(new byte[] { 1, 0 });
        }

        private byte[] GetMetaDataBytes()
        {
            var bw = new BinaryWriter(new MemoryStream());
            bw.Write(0); //Version
            var mdBytes = Encoding.UTF8.GetBytes(MetadataXml.OuterXml);
            bw.Write(mdBytes.Length);
            bw.Write(mdBytes);
            if(MetadataContentPackage == null)
            {
                bw.Write(0x16);
                bw.Write(new byte[] { 0x50, 0x4b, 0x05, 0x06, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0 });
            }
            else
            {
                bw.Write(MetadataContentPackage.Length);
                bw.Write(MetadataContentPackage);
            }
            bw.Flush();            
            return ((MemoryStream)bw.BaseStream).ToArray();
        }

        internal ZipPackage PQPackage { get; set; }
        internal byte[] MetadataContentPackage { get; set; }
        /// <summary>        
        /// <para>The plain-text document that contains the Power Query Formula for each query in the spreadsheet.</para>
        /// <para>It is fully specified by <see href="https://learn.microsoft.com/en-us/powerquery-m/power-query-m-language-specification">[MSDOCS - MLANG]</see> except for the following special rules:</para>
        /// <list type="bullet">
        /// <item><description>Only a single section is allowed and MUST be named “Section1”.</description></item>
        /// <item><description>Section member names MUST NOT contain periods, double quotes, tabs, leading/trailing whitespace, or line/carriage returns.Also, they MUST NOT be blank or consist only of whitespace</description></item>
        /// <item><description>All section members SHOULD be shared.</description></item>
        /// </list>
        /// </summary>        
        public string PowerQueryFormulas 
        { 
            get; 
            set; 
        }
        /// <summary>
        /// Permission setings for the Power Query connection.
        /// </summary>
        public XmlDocument PermissionsXml
        {
            get;
            private set;
        }
        /// <summary>
        /// Metadata settings for the Power Query connection.
        /// </summary>
        public XmlDocument MetadataXml
        {
            get;            
            private set;
        }
    }
}