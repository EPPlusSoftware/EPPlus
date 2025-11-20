using EPPlus.Fonts.OpenType.Scanner;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests
{
    [TestClass]
    public class SubsettingTests
    {
        private static string _fontFolder = string.Empty;
        private static List<string> _fontFolders = new List<string>();

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            _fontFolder = Path.Combine(AppContext.BaseDirectory, "Fonts");
            _fontFolders.Clear();
            _fontFolders.Add(_fontFolder);
            OpenTypeFonts.ClearFontCache();
        }

        [TestMethod]
        public void SubsetRoboto()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var fullFontBytes = font.Serialize();
            var subsetFont = font.CreateSubset(new char[] { 'A', 'B', 'C' });
            var subsetFontBytes = subsetFont.Serialize();
            if (File.Exists(@"c:\Temp\SubsetFontRoboto2.otf")) File.Delete(@"c:\Temp\SubsetFontRoboto2.otf");
            File.WriteAllBytes(@"c:\Temp\SubsetFontRoboto2.otf", subsetFontBytes);
        }

        [TestMethod]
        public void SubsetRobotoRead()
        {
            var fontFolder = @"c:\Temp\";
            var fontFolders = new List<string> { fontFolder };
            var font = OpenTypeFonts.GetFontData(fontFolders, "Roboto", "Regular", false);
            var fullFontBytes = font.Serialize();

            var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            font.CmapTable.Serialize(writer);
            var cmapBytes = ms.ToArray();
            
        }
    }
}
