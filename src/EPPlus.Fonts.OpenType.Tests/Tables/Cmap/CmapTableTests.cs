using EPPlus.Fonts.OpenType.Tests.Helpers;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests.Tables
{
    [TestClass]
    public class CmapTableTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        [DataRow("Mulish-Regular.ttf")]
        [DataRow("Roboto-Regular.ttf")]
        [DataRow("OpenSans-Regular.ttf")]
        [DataRow("EBGaramond-Regular.ttf")]
        [DataRow("CrimsonText-Regular.ttf")]
        public void Cmap_Subset_EncodingRecordsAreSorted(string fileName)
        {
            var font = new OpenTypeFont(File.ReadAllBytes(Path.Combine(FontFolder, fileName)));
            var bytes = new OpenTypeFontSerializer(font.CreateSubset("Hello World 123")).Serialize();
            SaveFontForCurrentTest(new OpenTypeFont(bytes), Path.GetFileNameWithoutExtension(fileName));

            FontTestHelper.FindTableInDirectory(bytes, "cmap", out int cmapOffset, out _);
            int numTables = (bytes[cmapOffset + 2] << 8) | bytes[cmapOffset + 3];

            var records = new List<KeyValuePair<int, int>>();
            for (int i = 0; i < numTables; i++)
            {
                int recordStart = cmapOffset + 4 + i * 8;
                int platformId = (bytes[recordStart] << 8) | bytes[recordStart + 1];
                int encodingId = (bytes[recordStart + 2] << 8) | bytes[recordStart + 3];
                records.Add(new KeyValuePair<int, int>(platformId, encodingId));
            }

            Console.WriteLine("Encoding records: " + string.Join(", ", records.Select(r => "(" + r.Key + "," + r.Value + ")").ToArray()));

            for (int i = 1; i < records.Count; i++)
            {
                bool inOrder = records[i - 1].Key < records[i].Key
                    || (records[i - 1].Key == records[i].Key && records[i - 1].Value < records[i].Value);
                Assert.IsTrue(inOrder, "Encoding records are not sorted by platformID, encodingID");
            }
        }
    }
}