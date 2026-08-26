using EPPlus.Fonts.OpenType.Scanner;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Fonts.OpenType.Tests.FontScanning
{
    [TestClass]
    public class ArialBlackTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void ScanArialBlack_ShouldReturnArialBlack()
        {
            var face = FontScannerV2.FindBestMatch(string.Empty, "Arial Black", FontSubFamily.Regular, true);
            if(face == null)
            {
                Assert.Inconclusive();
            }
            Assert.AreEqual("Arial Black", face.FamilyName, $"face.FamilyName was not 'Arial Black' as expected but '{face.FamilyName}'");
            Assert.IsTrue(face.IsExactMatch, "face.IsExactMatch was false");
        }

        [TestMethod]
        public void LoadArialBlackFullFont_ShouldReturnArialBlack()
        {
            var factory = new OpenTypeFontEngine();
            var availability = factory.GetFontAvailability("Arial Black");
            if(availability == FontAvailability.NotFound)
            {
                Assert.Inconclusive();
            }
            var font = factory.LoadFont("Arial Black");
            Assert.IsNotNull(font);
            Assert.AreEqual("Arial Black", font.FullName);
        }

        [TestMethod]
        public void Dump_AllFacesNamedLikeArialBlack()
        {
            var directories = System.Array.Empty<string>();
            var allFaces = FontScannerV2.EnumerateAllFaces(
                EPPlus.Fonts.OpenType.FontResolver.DefaultFontLocations.GetLocationsCollection(
                    directories, searchSystemDirectories: true));

            bool foundAny = false;
            foreach (var face in allFaces)
            {
                if (face.FamilyName != null &&
                    face.FamilyName.IndexOf("black", System.StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    foundAny = true;
                    Console.WriteLine(
                        "FamilyName='{0}'  SubfamilyName='{1}'  Subfamily={2}  FsSelection=0x{3:X4}  FilePath={4}",
                        face.FamilyName, face.SubfamilyName, face.Subfamily, face.FsSelection, face.FilePath);
                }
            }

            Assert.IsTrue(foundAny, "No installed face with 'black' in the family name was found at all.");
        }
    }
}
