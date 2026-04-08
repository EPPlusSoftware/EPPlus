using BenchmarkDotNet.Attributes;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Fonts;

[MemoryDiagnoser]
[SimpleJob(warmupCount: 1, iterationCount: 3)]
public class ExtractCharWidthsBenchmark
{
    private ITextShaper _shaper;
    private string _shortText;
    private string _mediumText;
    private string _longText;
    private ShapingOptions _options;

    [GlobalSetup]
    public void Setup()
    {
        var fontFolders = new List<string> { /* your font paths */ };
        var font = OpenTypeFonts.LoadFont("Calibri");
        _shaper = new TextShaper(font);
        _options = ShapingOptions.Default;

        // Short: typical Excel cell
        _shortText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit."; // 56 chars

        // Medium: single paragraph
        _mediumText = new string('x', 550); // Simulate 550 char paragraph

        // Long: full 20 paragraphs
        _longText = new string('x', 11000); // Simulate 11k chars
    }

    [Benchmark]
    public double[] ExtractCharWidths_Short()
    {
        return _shaper.ExtractCharWidths(_shortText, 11f, _options);
    }

    [Benchmark]
    public double[] ExtractCharWidths_Medium()
    {
        return _shaper.ExtractCharWidths(_mediumText, 11f, _options);
    }

    [Benchmark]
    public double[] ExtractCharWidths_Long()
    {
        return _shaper.ExtractCharWidths(_longText, 11f, _options);
    }

    // For comparison: what does Shape() alone allocate?
    [Benchmark]
    public ShapedText ShapeOnly_Short()
    {
        return _shaper.Shape(_shortText, _options);
    }

    [Benchmark]
    public ShapedText ShapeOnly_Medium()
    {
        return _shaper.Shape(_mediumText, _options);
    }

    [Benchmark]
    public ShapedText ShapeOnly_Long()
    {
        return _shaper.Shape(_longText, _options);
    }
}