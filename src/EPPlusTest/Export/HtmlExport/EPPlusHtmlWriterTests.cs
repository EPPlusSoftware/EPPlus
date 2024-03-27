using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Export.HtmlExport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace EPPlusTest.Export.HtmlExport
{
    [TestClass]
    public class EPPlusHtmlWriterTests
    {
        [TestMethod]
        public void ShouldWriteTable()
        {
            using(var ms = new MemoryStream())
            {
                var writer = new HtmlWriter(ms, Encoding.UTF8);
                writer.RenderBeginTagAsync(HtmlElements.Table).Wait();
                writer.RenderEndTagAsync(HtmlElements.Table).Wait();
                var reader = new StreamReader(ms);
                ms.Position = 0;
                var result = reader.ReadToEnd();
                Assert.AreEqual("<table></table>", result);
            }
        }

        [TestMethod]
        public void ShouldWriteTableWithClass()
        {
            using (var ms = new MemoryStream())
            {
                var writer = new HtmlWriter(ms, Encoding.UTF8);

                var attributes = new List<EpplusHtmlAttribute> { new EpplusHtmlAttribute() };
                attributes[0].AttributeName = HtmlAttributes.Class;
                attributes[0].Value = "myClass";

                writer.RenderBeginTagAsync( HtmlElements.Table, attributes).Wait();
                writer.RenderEndTagAsync(HtmlElements.Table).Wait();
                var reader = new StreamReader(ms);
                ms.Position = 0;
                var result = reader.ReadToEnd();
                Assert.AreEqual("<table class=\"myClass\"></table>", result);
            }
        }

        [TestMethod]
        public void ShouldWriteLinkWithHrefAndTarget()
        {
            using (var ms = new MemoryStream())
            {
                var writer = new HtmlWriter(ms, Encoding.UTF8);

                var attributes = new List<EpplusHtmlAttribute>
                {
                    new EpplusHtmlAttribute { AttributeName = HtmlAttributes.Href, Value = "http://epplussoftware.com" },
                    new EpplusHtmlAttribute { AttributeName = HtmlAttributes.Target, Value = "_blank" }
                };

                writer.RenderBeginTagAsync(HtmlElements.A, attributes).Wait();
                writer.WriteAsync("EPPlus Software").Wait();
                writer.RenderEndTagAsync(HtmlElements.A).Wait();
                var reader = new StreamReader(ms);
                ms.Position = 0;
                var result = reader.ReadToEnd();
                Assert.AreEqual("<a href=\"http://epplussoftware.com\" target=\"_blank\">EPPlus Software</a>", result);
            }
        }

        [TestMethod]
        public void ShouldWriteTableWithNestedElements()
        {
            using (var ms = new MemoryStream())
            {
                var writer = new HtmlWriter(ms, Encoding.UTF8);
                writer.RenderBeginTagAsync(HtmlElements.Table).Wait();
                writer.RenderBeginTagAsync(HtmlElements.Thead).Wait();
                writer.RenderBeginTagAsync(HtmlElements.TableRow).Wait();
                writer.RenderBeginTagAsync(HtmlElements.TableHeader).Wait();
                writer.WriteAsync("test1").Wait();
                writer.RenderEndTagAsync(HtmlElements.TableHeader).Wait();
                writer.RenderBeginTagAsync(HtmlElements.TableHeader).Wait();
                writer.WriteAsync("test2").Wait();
                writer.RenderEndTagAsync(HtmlElements.TableHeader).Wait();
                writer.RenderEndTagAsync(HtmlElements.TableRow).Wait();
                writer.RenderEndTagAsync(HtmlElements.Thead).Wait();
                writer.RenderEndTagAsync(HtmlElements.Table).Wait();
                var reader = new StreamReader(ms);
                ms.Position = 0;
                var result = reader.ReadToEnd();
                Assert.AreEqual("<table><thead><tr><th>test1</th><th>test2</th></tr></thead></table>", result);
            }
        }

        [TestMethod]
        public void ShouldWriteTableWithNestedElementsAndIndent()
        {
            using (var ms = new MemoryStream())
            {
                var writer = new HtmlWriter(ms, Encoding.UTF8);
                writer.RenderBeginTagAsync(HtmlElements.Table).Wait();
                writer.Indent++;
                writer.WriteLineAsync().Wait();
                writer.RenderBeginTagAsync(HtmlElements.Thead).Wait();
                writer.Indent++;
                writer.WriteLineAsync().Wait();
                writer.RenderBeginTagAsync(HtmlElements.TableRow).Wait();
                writer.RenderEndTagAsync(HtmlElements.TableRow).Wait();
                writer.Indent--;
                writer.WriteLineAsync().Wait();
                writer.RenderEndTagAsync(HtmlElements.Thead).Wait();
                writer.Indent--;
                writer.WriteLineAsync().Wait();
                writer.RenderEndTagAsync(HtmlElements.Table).Wait();
                var reader = new StreamReader(ms);
                ms.Position = 0;
                var result = reader.ReadToEnd();
                Assert.AreEqual($"<table>{Environment.NewLine}  <thead>{Environment.NewLine}    <tr></tr>{Environment.NewLine}  </thead>{Environment.NewLine}</table>", result);
            }
        }
    }
}
