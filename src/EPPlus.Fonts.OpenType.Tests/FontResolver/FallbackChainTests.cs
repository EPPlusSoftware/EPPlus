using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Tests.FontResolver
{
    /// <summary>
    /// Tests for <see cref="OpenTypeFontEngine.GetFallbackChain"/>: the user-configured chain
    /// followed by the built-in chain, without duplicates, in resolution order.
    /// </summary>
    [TestClass]
    public class FallbackChainTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private static OpenTypeFontEngine CreateEngine(Action<IEpplusFontConfiguration>? configure = null)
        {
            return new OpenTypeFontEngine(cfg =>
            {
                cfg.SearchSystemDirectories = false;
                configure?.Invoke(cfg);
            });
        }

        [TestMethod]
        public void BuiltinChain_AptosNarrow()
        {
            using (var engine = CreateEngine())
            {
                CollectionAssert.AreEqual(
                    new[] { "Liberation Sans Narrow", "Arial Narrow", "Carlito", "Liberation Sans", "Arial" },
                    engine.GetFallbackChain("Aptos Narrow"));
            }
        }

        [TestMethod]
        public void BuiltinChain_LookupIsCaseInsensitive()
        {
            using (var engine = CreateEngine())
            {
                CollectionAssert.AreEqual(
                    engine.GetFallbackChain("Calibri"),
                    engine.GetFallbackChain("cALIBRI"));
                Assert.IsTrue(engine.GetFallbackChain("cALIBRI").Length > 0);
            }
        }

        [TestMethod]
        public void UserChain_ComesBeforeBuiltinChain()
        {
            using (var engine = CreateEngine(cfg => cfg.FontFallbacks["Calibri"] = new[] { "Roboto", "Open Sans" }))
            {
                CollectionAssert.AreEqual(
                    new[] { "Roboto", "Open Sans", "Carlito", "Liberation Sans", "Arial", "Helvetica" },
                    engine.GetFallbackChain("Calibri"));
            }
        }

        [TestMethod]
        public void Duplicates_AreRemoved_FirstOccurrenceWins()
        {
            // "carlito" from the user chain wins over "Carlito" from the built-in chain,
            // and the repeated "Roboto" in the user chain is dropped.
            using (var engine = CreateEngine(cfg => cfg.FontFallbacks["Calibri"] = new[] { "Roboto", "carlito", "ROBOTO" }))
            {
                CollectionAssert.AreEqual(
                    new[] { "Roboto", "carlito", "Liberation Sans", "Arial", "Helvetica" },
                    engine.GetFallbackChain("Calibri"));
            }
        }

        [TestMethod]
        public void FontItself_IsNotInChain()
        {
            using (var engine = CreateEngine(cfg => cfg.FontFallbacks["Calibri"] = new[] { "calibri", "Roboto" }))
            {
                var chain = engine.GetFallbackChain("Calibri");
                Assert.IsFalse(chain.Any(n => string.Equals(n, "Calibri", StringComparison.OrdinalIgnoreCase)));
                Assert.AreEqual("Roboto", chain[0]);
            }
        }

        [TestMethod]
        public void EmptyNamesInUserChain_AreSkipped()
        {
            using (var engine = CreateEngine(cfg => cfg.FontFallbacks["My Font"] = new[] { "", null!, "Roboto" }))
            {
                CollectionAssert.AreEqual(new[] { "Roboto" }, engine.GetFallbackChain("My Font"));
            }
        }

        [TestMethod]
        public void UnknownFont_GivesEmptyChain()
        {
            using (var engine = CreateEngine())
            {
                Assert.AreEqual(0, engine.GetFallbackChain("No Such Font 123").Length);
            }
        }

        [TestMethod]
        public void NullOrEmptyName_GivesEmptyChain()
        {
            using (var engine = CreateEngine())
            {
                Assert.AreEqual(0, engine.GetFallbackChain(null!).Length);
                Assert.AreEqual(0, engine.GetFallbackChain(string.Empty).Length);
            }
        }

        [TestMethod]
        public void LastResortFont_IsNotInChain()
        {
            using (var engine = CreateEngine())
            {
                Assert.IsFalse(engine.GetFallbackChain("Aptos Narrow")
                    .Any(n => string.Equals(n, "Archivo Narrow", StringComparison.OrdinalIgnoreCase)));
            }
        }
    }
}
