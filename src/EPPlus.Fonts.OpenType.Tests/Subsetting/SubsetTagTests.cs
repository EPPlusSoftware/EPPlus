using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Name;
using EPPlus.Fonts.OpenType.Tests;
using OfficeOpenXml.Interfaces.Fonts;
using System.Text.RegularExpressions;

[TestClass]
public class SubsetTagTests : FontTestBase
{
    public override TestContext? TestContext { get; set; }

    private static NameTable CreateNameTable()
    {
        Func<ushort, ushort, ushort, NameRecordTypes, string, NameRecord> rec =
            (platform, encoding, lang, type, text) => new NameRecord
            {
                platformId = platform,
                encodingId = encoding,
                languageID = lang,
                nameId = (ushort)type,
                RecordType = type,
                Name = text
            };

        return new NameTable
        {
            NameRecords = new[]
            {
                rec(3, 1, 0x0409, NameRecordTypes.FontFamilyName, "Calibri"),
                rec(3, 1, 0x0409, NameRecordTypes.FontSubfamilyName, "Bold"),
                rec(3, 1, 0x0409, NameRecordTypes.FullFontName, "Calibri Bold"),
                rec(3, 1, 0x0409, NameRecordTypes.PostScriptName, "Calibri-Bold"),
                rec(1, 0, 0, NameRecordTypes.FullFontName, "Calibri Bold"),
                rec(1, 0, 0, NameRecordTypes.PostScriptName, "Calibri-Bold")
            }
        };
    }

    [TestMethod]
    public void ApplySubsetTag_PrefixesFullNameAndPostScriptName_OnAllPlatforms()
    {
        var nt = CreateNameTable();
        nt.ApplySubsetTag("ABCDEF+");

        foreach (var r in nt.NameRecords)
        {
            if (r.RecordType == NameRecordTypes.FullFontName || r.RecordType == NameRecordTypes.PostScriptName)
                Assert.IsTrue(r.Name.StartsWith("ABCDEF+"), r.Name);
        }
        Assert.AreEqual("ABCDEF+Calibri-Bold", nt.PostScriptName);
        Assert.AreEqual("ABCDEF+Calibri Bold", nt.GetFullFontName());
    }

    [TestMethod]
    public void ApplySubsetTag_DoesNotTouchFamilyOrSubfamily()
    {
        var nt = CreateNameTable();
        nt.ApplySubsetTag("ABCDEF+");

        Assert.AreEqual("Calibri", nt.NameRecords[0].Name);
        Assert.AreEqual("Bold", nt.NameRecords[1].Name);
    }

    [TestMethod]
    public void ApplySubsetTag_IsIdempotent()
    {
        var nt = CreateNameTable();
        nt.ApplySubsetTag("ABCDEF+");
        nt.ApplySubsetTag("ABCDEF+");
        Assert.AreEqual("ABCDEF+Calibri-Bold", nt.PostScriptName);
    }

    [TestMethod]
    public void Create_IsDeterministic_SixUppercaseLettersAndPlus()
    {
        var id = new FontKey("Calibri", FontSubFamily.Bold);
        var a = SubsetTag.Create(id, new HashSet<int> { 65, 66, 67 });
        var b = SubsetTag.Create(id, new HashSet<int> { 67, 66, 65 });
        Assert.AreEqual(a, b);
        Assert.IsTrue(System.Text.RegularExpressions.Regex.IsMatch(a, "^[A-Z]{6}\\+$"));
    }

    [TestMethod]
    public void Create_DiffersForDifferentCodePointsAndFonts()
    {
        var bold = new FontKey("Calibri", FontSubFamily.Bold);
        var regular = new FontKey("Calibri", FontSubFamily.Regular);
        var cps = new HashSet<int> { 65, 66 };
        Assert.AreNotEqual(SubsetTag.Create(bold, cps), SubsetTag.Create(bold, new HashSet<int> { 65 }));
        Assert.AreNotEqual(SubsetTag.Create(bold, cps), SubsetTag.Create(regular, cps));
    }

    [TestMethod]
    public void Subset_WithTag_SurvivesSerializeRoundtrip()
    {
        RequireFont(SystemFontsEngine, "Calibri", FontSubFamily.Regular);
        var font = SystemFontsEngine.LoadFont("Calibri");
        var subset = new SingleFontSubsetter().Subset(font, new HashSet<int> { 65, 66, 67 }, true);

        Assert.IsTrue(subset.IsSubset);
        Assert.IsTrue(Regex.IsMatch(subset.FullName, "^[A-Z]{6}\\+"), subset.FullName);

        var reloaded = new OpenTypeFont(subset.Serialize());
        Assert.AreEqual(subset.FullName, reloaded.FullName);
        Assert.AreEqual(subset.NameTable.PostScriptName, reloaded.NameTable.PostScriptName);
    }

    [TestMethod]
    public void Subset_WithTag_DoesNotMutateOriginalFont()
    {
        RequireFont(SystemFontsEngine, "Calibri", FontSubFamily.Regular);
        var font = SystemFontsEngine.LoadFont("Calibri");   
        var before = font.FullName;

        new SingleFontSubsetter().Subset(font, new HashSet<int> { 65 }, true);

        Assert.AreEqual(before, font.FullName);
        Assert.IsFalse(Regex.IsMatch(font.FullName, "^[A-Z]{6}\\+"));
    }
}