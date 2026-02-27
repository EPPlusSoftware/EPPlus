using BenchmarkDotNet.Attributes;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml.Interfaces.Fonts;

/// <summary>
/// Benchmarks for repeated cache clearing scenarios
/// </summary>
[MemoryDiagnoser]
[SimpleJob(warmupCount: 3, iterationCount: 5)]
public class FontCacheClearingBenchmarks
{
    private List<string> _fontFolders;

    [GlobalSetup]
    public void Setup()
    {
        var fontsPath = Path.Combine(System.AppContext.BaseDirectory, "Fonts");

        if (!Directory.Exists(fontsPath))
        {
            throw new DirectoryNotFoundException($"Fonts directory not found: {fontsPath}");
        }

        _fontFolders = new List<string> { fontsPath };
    }

    [Benchmark]
    public OpenTypeFont Load_Clear_Load_Pattern()
    {
        // Simulates pattern where cache is cleared between operations
        OpenTypeFonts.ClearFontCache();
        var font1 = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular);

        OpenTypeFonts.ClearFontCache();
        var font2 = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular);

        return font2;
    }

    [Benchmark]
    public OpenTypeFont Load_Reuse_Pattern()
    {
        // Simulates pattern where cache is NOT cleared (optimal)
        var font1 = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular);
        var font2 = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular);

        return font2;
    }
}