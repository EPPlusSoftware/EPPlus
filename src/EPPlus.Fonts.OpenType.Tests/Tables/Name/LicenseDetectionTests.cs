using EPPlus.Fonts.OpenType.Tables.Name;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Tests.Tables
{
    /// <summary>
    /// Tests for licence detection in the name table: <see cref="NameTable.HasOpenLicense"/>,
    /// <see cref="NameTable.GetLicenseDescription"/> and <see cref="NameTable.GetLicenseInfoUrl"/>.
    /// </summary>
    [TestClass]
    public class LicenseDetectionTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private const string MicrosoftDescription =
            "Microsoft supplied font. You may use this font to create, display, and print content as permitted by the license terms or terms of use, of the Microsoft product, service, or content in which this font was included.";

        /// <summary>
        /// Loads a private copy of a test font, so changes to its name records do not leak into the
        /// shared engine.
        /// </summary>
        internal static OpenTypeFont LoadPrivate(string family)
        {
            RequireFont(TestFolderEngine, family);
            return TestFolderEngine.LoadFont(family, FontSubFamily.Regular, ignoreCache: true);
        }

        /// <summary>
        /// Replaces the licence records (nameID 13 and 14) of the font. Pass null to leave a record out.
        /// </summary>
        internal static void SetLicenseRecords(OpenTypeFont font, string description, string url)
        {
            var records = font.NameTable.NameRecords
                .Where(r => r.RecordType != NameRecordTypes.LicenseDescription && r.RecordType != NameRecordTypes.LicenseInfoUrl)
                .ToList();

            if (description != null)
                records.Add(CreateWindowsRecord(NameRecordTypes.LicenseDescription, description));
            if (url != null)
                records.Add(CreateWindowsRecord(NameRecordTypes.LicenseInfoUrl, url));

            font.NameTable.NameRecords = records.ToArray();
        }

        private static NameRecord CreateWindowsRecord(NameRecordTypes type, string text)
        {
            return new NameRecord
            {
                platformId = 3,
                encodingId = 1,
                languageID = 0x0409,
                nameId = (ushort)type,
                RecordType = type,
                Name = text
            };
        }

        // -----------------------------------------------------------------------------------------
        // Test fonts as shipped
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void Roboto_IsApache()
        {
            var font = LoadPrivate("Roboto");
            Assert.IsTrue(font.NameTable.HasOpenLicense());
            StringAssert.Contains(font.NameTable.GetLicenseDescription(), "Apache License");
            StringAssert.Contains(font.NameTable.GetLicenseInfoUrl(), "apache.org/licenses");
        }

        [TestMethod]
        [DataRow("Open Sans")]
        [DataRow("EB Garamond")]
        [DataRow("Mulish")]
        public void OflFont_IsOpen(string family)
        {
            var font = LoadPrivate(family);
            Assert.IsTrue(font.NameTable.HasOpenLicense());
            StringAssert.Contains(font.NameTable.GetLicenseDescription(), "SIL Open Font License");
        }

        // -----------------------------------------------------------------------------------------
        // Rules, on private copies with the licence records replaced
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void NoLicenseRecords_IsNotOpen()
        {
            var font = LoadPrivate("Roboto");
            SetLicenseRecords(font, null, null);

            Assert.IsFalse(font.NameTable.HasOpenLicense());
            Assert.IsNull(font.NameTable.GetLicenseDescription());
            Assert.IsNull(font.NameTable.GetLicenseInfoUrl());
        }

        [TestMethod]
        public void CommercialDescription_IsNotOpen()
        {
            var font = LoadPrivate("Roboto");
            SetLicenseRecords(font, MicrosoftDescription, "https://docs.microsoft.com/typography");

            Assert.IsFalse(font.NameTable.HasOpenLicense());
            Assert.AreEqual(MicrosoftDescription, font.NameTable.GetLicenseDescription());
        }

        [TestMethod]
        [DataRow("This Font Software is licensed under the SIL Open Font License, Version 1.1.")]
        [DataRow("Licensed under the Apache License, Version 2.0")]
        [DataRow("Copyright 2011 Canonical Ltd. Licensed under the Ubuntu Font Licence 1.0")]
        [DataRow("Fonts are (c) Bitstream (see below). DejaVu changes are in public domain. Bitstream Vera Fonts Copyright ...")]
        [DataRow("this font software is licensed under the sil open font license")] // case-insensitive
        public void OpenDescription_WithoutUrl_IsOpen(string description)
        {
            var font = LoadPrivate("Roboto");
            SetLicenseRecords(font, description, null);
            Assert.IsTrue(font.NameTable.HasOpenLicense());
        }

        [TestMethod]
        [DataRow("http://scripts.sil.org/OFL")]
        [DataRow("https://openfontlicense.org")]
        [DataRow("http://www.apache.org/licenses/LICENSE-2.0")]
        public void OpenUrl_WithoutDescription_IsOpen(string url)
        {
            var font = LoadPrivate("Roboto");
            SetLicenseRecords(font, null, url);
            Assert.IsTrue(font.NameTable.HasOpenLicense());
        }

        [TestMethod]
        public void OpenLicenseOnOtherPlatformRecord_IsOpen()
        {
            // Any record counts, not only the English Windows record.
            var font = LoadPrivate("Roboto");
            SetLicenseRecords(font, MicrosoftDescription, null);
            var records = font.NameTable.NameRecords.ToList();
            records.Add(new NameRecord
            {
                platformId = 1,
                encodingId = 0,
                languageID = 0,
                nameId = (ushort)NameRecordTypes.LicenseDescription,
                RecordType = NameRecordTypes.LicenseDescription,
                Name = "Licensed under the SIL Open Font License"
            });
            font.NameTable.NameRecords = records.ToArray();

            Assert.IsTrue(font.NameTable.HasOpenLicense());
        }

        [TestMethod]
        public void GplFontException_IsNotOpen()
        {
            // Open decision in the plan: GPLv2 with font exception (Liberation 1.x) is not accepted yet.
            var font = LoadPrivate("Roboto");
            SetLicenseRecords(font,
                "Licensed under the GNU General Public License v2 with the font exception",
                "http://www.gnu.org/licenses/gpl-2.0.html");
            Assert.IsFalse(font.NameTable.HasOpenLicense());
        }
    }
}