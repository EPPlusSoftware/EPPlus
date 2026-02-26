using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.Tests.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests.Serialization
{
    [TestClass]
    public class KernSerializationTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        [Ignore("Was only able to find fonts with kern table among Windows fonts. These cannot be distributed with the test project due to licensing.")]
        public void SerializeKernTable()
        {
            var ffi = FontScannerV2.FindBestMatch(@"c:\windows\fonts", "Arial", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("kern");

            var font = OpenTypeFonts.LoadFont("Arial");
            var kernBytes = font?.KernTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, kernBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, kernBytes);
        }

        [TestMethod, Ignore]
        public void FindKernTable()
        {
            var fonts = FontScannerV2.GetAllScannedFontsInPath(@"c:\windows\fonts");
            var kernFonts = new List<FontFaceInfo>();
            foreach (var font in fonts)
            {
                if (font.TableRecords.ContainsKey("kern") || font.TableRecords.ContainsKey("Kern"))
                {
                    kernFonts.Add(font);
                }
            }
            var c = kernFonts.Count;
        }
    }
}
