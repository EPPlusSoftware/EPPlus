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
using OfficeOpenXml.Constants;
using OfficeOpenXml.Data.CustomXml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;

namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// Settings for power query connections. These settings are read from the CustomXml in the package with the key DataMashup
    /// </summary>
    public class ExcelPowerQuerySettings
    {
        XmlNamespaceManager _nsm;
        internal ExcelPowerQuerySettings()
        {
            _nsm = CreateNsm();
        }

        private XmlNamespaceManager CreateNsm()
        {
            var nsm = new XmlNamespaceManager(new NameTable());
            nsm.AddNamespace("d", "http://schemas.microsoft.com/DataMashup");
            nsm.AddNamespace("xsd", "http://www.w3.org/2001/XMLSchema");
            nsm.AddNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance");
            return nsm;
        }


        internal ExcelPowerQuerySettings(byte[] blob)
        {
            _nsm = CreateNsm();

            using var ms = new MemoryStream(blob);
            var br = new BinaryReader(ms);
            var version = br.ReadInt32();
            var size = br.ReadInt32();
            var pck = br.ReadBytes((int)size);

            PQPackage = new ZipPackage(new MemoryStream(pck));
            ZipPackagePart section1MPart;

            section1MPart = PQPackage.GetPartByContentType(ContentTypes.contentTypeMLanguage);
            if(section1MPart==null)
            {
                section1MPart = PQPackage.GetPart(new Uri("/formulas/section1.m", UriKind.Relative));
            }
            var mms = section1MPart.GetStream();
            var reader = new StreamReader(mms, Encoding.UTF8);
            Formulas = reader.ReadToEnd();

            var configPart = PQPackage.GetPart(new Uri("/config/package.xml", UriKind.Relative));
            var cms = configPart.GetStream();

            reader = new StreamReader(cms, Encoding.UTF8);
            LoadPackageInfo(reader.ReadToEnd());
            
            size = br.ReadInt32();
            var permBytes = br.ReadBytes((int)size);
            var permissionXml = Encoding.UTF8.GetString(permBytes);
            LoadPermissions(permissionXml);
            
            size = br.ReadInt32();
            version = br.ReadInt32();
            size = br.ReadInt32();
            var metadataXml = Encoding.UTF8.GetString(br.ReadBytes((int)size));
            LoadMetadataXml(metadataXml);
            
            size = br.ReadInt32();            
            MetadataContentPackage = br.ReadBytes((int)size);
            size = br.ReadInt32();

            var packageBinding = br.ReadBytes((int)size);
            // Data protection (DPAPI) only works in windows as it's tied to the current user.
            //#if(!Core)            
            //            var protectedData = ProtectedData.Unprotect(packageBinding, UTF8Encoding.UTF8.Save("DataExplorer Package Components"), DataProtectionScope.CurrentUser);
            //            br = new BinaryReader(new MemoryStream(protectedData));
            //            size = br.ReadInt32();
            //            var hash1 = br.ReadBytes(size);
            //            size = br.ReadInt32();
            //            var hash2 = br.ReadBytes(size);

            //            var sha = SHA256.Create();
            //            var calcHash1 = sha.ComputeHash(pck);
            //            var calcHash2 = sha.ComputeHash(permBytes);
            //#endif
        }
        internal XmlHelper _permissionsXh;
        internal XmlHelper _packageInfoXh;
        internal XmlHelper _metaDataXh;
        private void LoadPackageInfo(string xml)
        {
            var xmlDoc = new XmlDocument();
            XmlHelper.LoadXmlSafe(xmlDoc, xml, Encoding.UTF8);
            _packageInfoXh = XmlHelperFactory.Create(_nsm, xmlDoc.DocumentElement);
            Version = _packageInfoXh.GetXmlNodeString("Version");
            MinimumVersion = _packageInfoXh.GetXmlNodeString("MinVersion");
            CultureCode = _packageInfoXh.GetXmlNodeString("Culture");
        }
        private void SavePackageInfo()
        {
            _packageInfoXh.SetXmlNodeString("Version", Version);
            _packageInfoXh.SetXmlNodeString("MinVersion", MinimumVersion);
            if (string.IsNullOrEmpty(CultureCode))
            {
                CultureCode = Thread.CurrentThread.CurrentCulture.Name;
            }
            _packageInfoXh.SetXmlNodeString("Culture", CultureCode);
        }
        private void LoadPermissions(string xml)
        {
            var xmlDoc = new XmlDocument();
            XmlHelper.LoadXmlSafe(xmlDoc, xml, Encoding.UTF8);
            _permissionsXh = XmlHelperFactory.Create(_nsm, xmlDoc.DocumentElement);
            Permissions = new ExcelPowerQueryPermissions();
            Permissions.CanEvaluateFuturePackages = _permissionsXh.GetXmlNodeBool("CanEvaluateFuturePackages");
            Permissions.FirewallEnabled = _permissionsXh.GetXmlNodeBool("FirewallEnabled");
            Permissions.PrivacyLevel = _permissionsXh.GetXmlEnum("WorkbookGroupType", ePowerQueryPermissionWorkbookGroupType.None);
        }
        private void SavePermissions()
        {
            _permissionsXh.SetXmlNodeBool("CanEvaluateFuturePackages", Permissions.CanEvaluateFuturePackages);
            _permissionsXh.SetXmlNodeBool("FirewallEnabled", Permissions.FirewallEnabled);
            _permissionsXh.SetXmlNodeString("WorkbookGroupType", Permissions.PrivacyLevel.ToEnumString(ePowerQueryPermissionWorkbookGroupType.None), true);
        }
        /// <summary>
        /// Loads meta data from a xml document formatted according to the MS-QDEFF docmument - section 2.5.1.
        /// See also <seealso cref="MetadataXml"/>
        /// </summary>
        /// <param name="xml">The xml document</param>
        /// <exception cref="ArgumentException">If the xml fails to load or does not contain any items.</exception>
        public void LoadMetadataXml(string xml)
        {
            MetadataXml = new XmlDocument();
            try
            {
                XmlHelper.LoadXmlSafe(MetadataXml, xml, Encoding.UTF8);
            }
            catch(Exception ex)
            {
                throw new ArgumentException("The xml supplied failed to load. See inner exception for more details", ex);
            }

            _metaDataXh = XmlHelperFactory.Create(_nsm, MetadataXml.DocumentElement);
            var culture = new CultureInfo(CultureCode);
            MetadataItems.Clear();
            foreach (XmlNode n in _metaDataXh.GetNodes("Items/Item"))
            {
                MetadataItems.Add(new ExcelPowerQueryMetadataItem(_nsm, n, culture));
            }
            if (MetadataItems.Count == 0)
            {                
                LoadMetadataXml(defaultMetadataXml);
                throw (new ArgumentException("The meta data xml must contain at least one item."));
            }
        }
        private void SaveMetaData()
        {
            var itemsNode = _metaDataXh.GetNode("Items");
            itemsNode.InnerXml = "";
            var culture = new CultureInfo(CultureCode);

            foreach (var item in MetadataItems)
            {
                var itemNode = _metaDataXh.CreateNode("Items/Item", false, true);
                var itemXh = XmlHelperFactory.Create(_nsm, itemNode);
                itemXh.SetXmlNodeString("ItemLocation/ItemType", item.ItemType.ToString());
                itemXh.SetXmlNodeString("ItemLocation/ItemPath", item.ItemPath);

                itemXh.CreateNode("StableEntries");
                foreach(var entry in item.Entries)
                {
                    XmlElement e = (XmlElement)itemXh.CreateNode("StableEntries/Entry", false, true);
                    e.SetAttribute("Type", entry.EntryType);
                    e.SetAttribute("Value", entry.GetValueAsText(culture));
                }
            }
        }

        internal void Save(ExcelCustomXmlCollection customXml)
        {
            var cx = customXml.FirstOrDefault(x => x.SchemasReferences.Contains(Schemas.schemaDataMashup));

            if (cx != null && Exists == false)
            {
                customXml.Remove(cx);
                return;
            }

            SavePackageInfo();
            SavePermissions();
            SaveMetaData();
            ZipPackagePart sectionMPart;
            if(PQPackage==null)
            {
                PQPackage=new ZipPackage(new MemoryStream());
                var pp = PQPackage.CreatePart(new Uri("/Config/Package.xml", UriKind.Relative), "text/xml");
                var sw = new StreamWriter(pp.GetStream(), new UTF8Encoding()); 
                sw.Write($"<?xml version=\"1.0\" encoding=\"utf-8\"?><Package xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><Version>2.147.503.0</Version><MinVersion>2.21.0.0</MinVersion><Culture>{Thread.CurrentThread.CurrentCulture.Name}</Culture></Package>");
                sw.Flush();
                sectionMPart = PQPackage.CreatePart(new Uri("/Formulas/Section1.m", UriKind.Relative), "application/x-ms-m");
            }
            else
            {
                sectionMPart = PQPackage.GetPartByContentType("application/x-ms-m");
                if(sectionMPart==null)
                {
                    var uri = new Uri("/Formulas/Section1.m", UriKind.Relative);
                    if(PQPackage.PartExists(uri))
                    {
                        sectionMPart = PQPackage.GetPart(uri);
                    }
                    else
                    {
                        sectionMPart = PQPackage.CreatePart(uri, "application/x-ms-m");
                    }
                }
            }

            var ms = sectionMPart.GetStream();
            var streamwriter = new StreamWriter(ms, new UTF8Encoding(false));
            streamwriter.Write(Formulas);
            streamwriter.Flush();
            //streamwriter.Close();

            var packageStream = new MemoryStream();
            PQPackage.Save(packageStream);
            var pckbytes = packageStream.ToArray();

            using var retMs = new MemoryStream();
            var bw= new BinaryWriter(retMs);
            bw.Write(0);
            bw.Write(pckbytes.Length);
            bw.Write(pckbytes);
            var permBytes = Encoding.UTF8.GetBytes(_permissionsXh.TopNode.OwnerDocument.OuterXml);
            bw.Write(permBytes.Length);
            bw.Write(permBytes);

            var metadataBytes = GetMetaDataBytes();
            bw.Write(metadataBytes.Length);
            bw.Write(metadataBytes);

            // Permissions binding. We set it to empty as DPAPI is only available on Windows.
            bw.Write(1);
            bw.Write((byte)0);
            bw.Flush();

            if (cx == null)
            {
                cx = new ExcelCustomXml() { SchemasReferences = { Schemas.schemaDataMashup}, CustomXml = new XmlDocument()  };
            }

            cx.CustomXml.LoadXml($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><DataMashup xmlns=\"{Schemas.schemaDataMashup}\">{Convert.ToBase64String(retMs.ToArray())}</DataMashup>");
            if (customXml.Contains(cx) == false)
            {
                customXml.Add(cx);
            }
        }
        private byte[] GetMetaDataBytes()
        {
            var bw = new BinaryWriter(new MemoryStream());
            bw.Write(0); //Version
            var mdBytes = Encoding.UTF8.GetBytes(_metaDataXh.TopNode.OwnerDocument.OuterXml);
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
        public string Formulas
        { 
            get; 
            set; 
        }
        internal string Version { get; set; }
        internal string MinimumVersion { get; set; }
        /// <summary>
        /// The culture code used to set parse numbers and dates
        /// </summary>
        public string CultureCode {  get; set; }
        /// <summary>
        /// Permission settings
        /// </summary>
        public ExcelPowerQueryPermissions Permissions
        {
            get;
            private set;
        } = null;

        /// <summary>
        /// Metadata settings for the Power Query connection. See MS-QDEFF - 2.5.1
        /// </summary>
        public XmlDocument MetadataXml
        {
            get;
            private set;
        }
        /// <summary>
        /// A collection of meta data items to describe the power query formulas.
        /// </summary>
        public List<ExcelPowerQueryMetadataItem> MetadataItems { get; } = new List<ExcelPowerQueryMetadataItem>();
        /// <summary>
        /// If any power query settings exists in the package. 
        /// <seealso cref="Create()"/>"/>
        /// </summary>
        public bool Exists
        {
            get
            {
                return Permissions != null;
            }
        }
        const string defaultPackageInfoXml = "<?xml version=\"1.0\" encoding=\"utf-8\"?><Package xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><Version>2.147.503.0</Version><MinVersion>2.21.0.0</MinVersion><Culture>{0}</Culture></Package>";
        const string defaultPermissionXml = "<?xml version=\"1.0\" encoding=\"utf-8\"?><PermissionList xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><CanEvaluateFuturePackages>false</CanEvaluateFuturePackages><FirewallEnabled>true</FirewallEnabled></PermissionList>";
        const string defaultMetadataXml = "<?xml version=\"1.0\" encoding=\"utf-8\"?><LocalPackageMetadataFile xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><Items><Item><ItemLocation><ItemType>AllFormulas</ItemType><ItemPath /></ItemLocation><StableEntries><Entry Type=\"Relationships\" Value=\"sAAAAAA==\" /></StableEntries></Item></Items></LocalPackageMetadataFile>";
        /// <summary>
        /// Create the DataMashup xml used for power query setting in the CustomXml. 
        /// This will initialize the PermissionsXml, MetadataXml and Formulas properties with empty settings.
        /// This enables the possibility to add power query connections to the workbook (provider=Microsoft.Mashup.OleDb.1). 
        /// Power query connections requires the M formula to be set in the <see cref="Formulas"/> under Section1 and meta data to be created in the <see cref="MetadataXml"/>.
        /// </summary>
        public void Create()
        {
            if(Permissions!=null)
            {
                throw (new InvalidOperationException("Power query settings already exist in the package."));
            }
            LoadPackageInfo(string.Format(defaultPackageInfoXml, Thread.CurrentThread.CurrentCulture.Name));
            LoadPermissions(defaultPermissionXml);
            LoadMetadataXml(defaultMetadataXml);
            Formulas = "section Section1;\n";
        }
    }
}