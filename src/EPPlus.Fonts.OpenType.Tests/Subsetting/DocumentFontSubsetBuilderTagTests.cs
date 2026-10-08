using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tests;
using OfficeOpenXml.Interfaces.Fonts;
using System.Text.RegularExpressions;

[TestClass]
public class DocumentFontSubsetBuilderTagTests : FontTestBase   // adjust to your base class
{
    private static readonly Regex TaggedName = new Regex("^[A-Z]{6}\\+.+$");

    public override TestContext? TestContext { get; set; }

    private DocumentFontSubsetBuilder CreateBuilder()
    {
        return new DocumentFontSubsetBuilder(SystemFontsEngine);
    }

    [TestMethod]
    public void Build_SubsettedFont_HasSubsetTag()
    {
        RequireFont(SystemFontsEngine, "Calibri", FontSubFamily.Regular);
        var builder = CreateBuilder();
        builder.AddText("Calibri", FontSubFamily.Regular, "ABC");
        builder.Build();

        var fonts = builder.GetFontsToEmbed().ToList();

        Assert.AreEqual(1, fonts.Count);
        // NOTE: property name for the font on SubsettedFont is assumed to be "Font".
        Assert.IsTrue(TaggedName.IsMatch(fonts[0].Font.FullName), fonts[0].Font.FullName);
    }

    [TestMethod]
    public void Build_TaggedSubset_IsStillFoundByOriginalIdentity()
    {
        // Regression guard: the tag must not leak into the identity used as dictionary key.
        RequireFont(SystemFontsEngine, "Calibri", FontSubFamily.Regular);
        var builder = CreateBuilder();
        builder.AddText("Calibri", FontSubFamily.Regular, "ABC");
        builder.Build();

        var provider = builder.GetShapingProvider("Calibri", FontSubFamily.Regular);
        Assert.IsNotNull(provider);

        var fonts = builder.GetFontsToEmbed().ToList();
        Assert.AreEqual("Calibri", fonts[0].Family);
        Assert.AreEqual(FontSubFamily.Regular, fonts[0].SubFamily);
    }

    [TestMethod]
    public void Build_ShapingProvider_ResolvesGlyphsFromTaggedSubset()
    {
        RequireFont(SystemFontsEngine, "Calibri", FontSubFamily.Regular);
        var builder = CreateBuilder();
        builder.AddText("Calibri", FontSubFamily.Regular, "ABC");
        builder.Build();

        var provider = builder.GetShapingProvider("Calibri", FontSubFamily.Regular);

        OpenTypeFont dest;
        ushort glyphId;
        provider.TryGetGlyphFont((uint)'A', out dest, out glyphId);

        Assert.AreNotEqual((ushort)0, glyphId);
        Assert.IsTrue(TaggedName.IsMatch(dest.FullName), dest.FullName);
    }

    [TestMethod]
    public void Build_DoesNotMutateCachedOriginalFont()
    {
        RequireFont(SystemFontsEngine, "Calibri", FontSubFamily.Regular);
        var before = SystemFontsEngine.LoadFont("Calibri").FullName;

        var builder = CreateBuilder();
        builder.AddText("Calibri", FontSubFamily.Regular, "ABC");
        builder.Build();

        var after = SystemFontsEngine.LoadFont("Calibri").FullName;
        Assert.AreEqual(before, after);
        Assert.IsFalse(TaggedName.IsMatch(after));
    }

    [TestMethod]
    public void Build_DifferentFontsAndSameInputTwice_TagsAreDistinctAndDeterministic()
    {
        RequireFont(SystemFontsEngine, "Calibri", FontSubFamily.Regular);
        RequireFont(SystemFontsEngine, "Calibri", FontSubFamily.Bold);

        var a = CreateBuilder();
        a.AddText("Calibri", FontSubFamily.Regular, "ABC");
        a.AddText("Calibri", FontSubFamily.Bold, "ABC");
        a.Build();

        var b = CreateBuilder();
        b.AddText("Calibri", FontSubFamily.Regular, "ABC");
        b.AddText("Calibri", FontSubFamily.Bold, "ABC");
        b.Build();

        var namesA = a.GetFontsToEmbed().Select(f => f.Font.FullName).OrderBy(n => n).ToList();
        var namesB = b.GetFontsToEmbed().Select(f => f.Font.FullName).OrderBy(n => n).ToList();

        Assert.AreEqual(2, namesA.Count);
        Assert.AreNotEqual(namesA[0].Substring(0, 6), namesA[1].Substring(0, 6));
        CollectionAssert.AreEqual(namesA, namesB);   // deterministic across builders
    }
}