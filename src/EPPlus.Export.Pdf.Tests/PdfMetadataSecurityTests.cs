using EPPlus.Export.Pdf.DocumentObjects;
using EPPlus.Export.Pdf.Settings;

namespace EPPlusTest.PDF
{
    [TestClass]
    public class PdfMetadataSecurityTests
    {
        [TestMethod]
        public void InfoDictionary_EscapesMetadata_NoPdfInjection()
        {
            // A parenthesis breakout that tries to close the string + dictionary and inject an action.
            var settings = new PdfDocumentSettings
            {
                Title = "Report) >> endobj /JavaScript (app.alert\\('pwn'\\)) (",
                Author = "a) b ( c \\ d",
                LicenseType = "Lic)Type",
                LicenseHolder = "Holder) ("
            };

            string body = new PdfInfoObject(1, settings).RenderBody();

            // Title/Author are hex strings now, not literals -> the payload cannot appear as tokens.
            StringAssert.Contains(body, "/Title <FEFF", "Title should be emitted as a hex string.");
            StringAssert.Contains(body, "/Author <FEFF", "Author should be emitted as a hex string.");

            // The injected keyword must not survive as literal PDF syntax anywhere in the object.
            Assert.IsFalse(body.Contains("/JavaScript"), "Injected /JavaScript token leaked into the Info dictionary.");
            Assert.IsFalse(body.Contains("endobj"), "Injected 'endobj' leaked into the Info dictionary.");

            // The Info dictionary must still be exactly one well-formed object:
            // one opening "<<" and one closing ">>", and no stray literal "()" breakout.
            Assert.AreEqual(1, CountOccurrences(body, "<<"), "Info dict should open exactly once.");
            Assert.AreEqual(1, CountOccurrences(body, ">>"), "Info dict should close exactly once.");
        }

        private static int CountOccurrences(string s, string sub)
        {
            int n = 0, i = 0;
            while ((i = s.IndexOf(sub, i, System.StringComparison.Ordinal)) >= 0) { n++; i += sub.Length; }
            return n;
        }
    }
}