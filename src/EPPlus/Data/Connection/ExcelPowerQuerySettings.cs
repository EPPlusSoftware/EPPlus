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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.Packaging;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using static System.Collections.Specialized.BitVector32;

namespace OfficeOpenXml.Data.Connection
{
    public class ExcelPowerQuerySettings
    {
        public ExcelPowerQuerySettings(byte[] blob)
        {
            //var cd = new CompoundDocument(b);
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

            var pck2 = br.ReadBytes((int)size);
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
        internal ZipPackage PQPackage { get; set; }
        /// <summary>        
        /// <para>The plain-text document that contains the Power Query Formula for each query in the spreadsheet.</para>
        /// <para>It is fully specified by <see href="https://learn.microsoft.com/en-us/powerquery-m/power-query-m-language-specification">[MSDOCS - MLANG]</see> except for the following special rules:</para>
        /// <list type="bullet">
        /// <item><description>Only a single section is allowed and MUST be named “Section1”.</description></item>
        /// <item><description>Section member names MUST NOT contain periods, double quotes, tabs, leading/trailing whitespace, or line/carriage returns.Also, they MUST NOT be blank or consist only of whitespace</description></item>
        /// <item><description>All section members SHOULD be shared.</description></item>
        /// </list>
        /// </summary>
        public string PowerQueryFormulas { get; set; }
        public XmlDocument PermissionsXml
        {
            get;
            private set;
        }
        public XmlDocument MetadataXml
        {
            get;
            
            private set;
        }
    }
}