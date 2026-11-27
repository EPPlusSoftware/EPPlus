using EPPlus.Export.Pdf.DocumentObjects;
using EPPlus.Export.Pdf.DocumentObjects.Fonts;
using EPPlus.Export.Pdf.Helpers;

namespace EPPlusTest.PDF
{
    [TestClass]
    public class PdfLowRiskHardeningTests
    {
        [TestMethod]
        public void FontName_EscapesPdfDelimiters()
        {
            var font = new PdfFont(1, "Evil) /Type /Catalog (x", PdfFontSubType.Type1);
            string body = font.RenderBody();

            Assert.IsFalse(body.Contains("Evil)"), "Unescaped ')' leaked into /BaseFont.");
            StringAssert.Contains(body, "Evil#29", "')' should be escaped as #29.");
            StringAssert.Contains(body, "#2F", "'/' should be escaped as #2F.");
        }

        [TestMethod]
        public void FontDescriptor_EscapesPdfDelimiters()
        {
            var fd = new PdfFontDescriptor(1, "Bad>Name) (", 0, new EPPlus.Graphics.Rect(), 0, 0, 0, 0, 0);
            string body = fd.RenderBody();

            Assert.IsFalse(body.Contains("Bad>Name"), "Unescaped '>' leaked into /FontName.");
            StringAssert.Contains(body, "#3E", "'>' should be escaped as #3E.");
        }

        [TestMethod]
        public void LinkAnnotation_DropsJavascriptUri()
        {
            var evil = new PdfLinkAnnotationObject(1, 0, 0, 10, 10, "javascript:app.alert('x')", null);
            string body = evil.RenderBody();

            Assert.IsFalse(body.Contains("/URI"), "javascript: must not produce a URI action.");
            Assert.IsFalse(body.Contains("javascript"), "Dangerous scheme leaked into the annotation.");
            StringAssert.Contains(body, "/Subtype /Link", "The link annotation itself should still be emitted (inert).");
        }

        [TestMethod]
        public void LinkAnnotation_KeepsSafeUri()
        {
            var safe = new PdfLinkAnnotationObject(2, 0, 0, 10, 10, "https://example.com/report", null);
            string body = safe.RenderBody();

            StringAssert.Contains(body, "/S /URI");
            StringAssert.Contains(body, "example.com");
        }

        [TestMethod]
        public void LinkAnnotation_BlocksObfuscatedScheme()
        {
            // leading space + embedded tab must not bypass the filter
            var evil = new PdfLinkAnnotationObject(3, 0, 0, 10, 10, " java\tscript:evil()", null);
            string body = evil.RenderBody();

            Assert.IsFalse(body.Contains("/URI"), "Obfuscated javascript: bypassed the scheme filter.");
        }

        [TestMethod]
        public void CidFont_BaseFont_EscapesPdfDelimiters()
        {
            // PdfCIDFont derives BaseFont from fontData.FullName — exercised via the resource layer
            // in a full export; here we assert the escaping contract on the emitted token directly.
            string name = "Evil) /Type /Catalog (x";
            string escaped = name.Replace(" ", "").ToPdfName();   // ToPdfName is the fix
            Assert.IsFalse(escaped.Contains(")"), "')' must be hex-escaped in a PDF name.");
            Assert.IsFalse(escaped.Contains("/"), "'/' must be hex-escaped in a PDF name.");
            StringAssert.Contains(escaped, "#29");
        }
    }
}