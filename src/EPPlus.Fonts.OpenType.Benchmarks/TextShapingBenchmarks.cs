using BenchmarkDotNet.Attributes;
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Benchmarks
{
    [MemoryDiagnoser]  // Shows memory allocations
    [SimpleJob(warmupCount: 3, iterationCount: 5)]  // 3 warmups, 5 measurements
    public class TextShapingBenchmarks
    {
        private OpenTypeFont _roboto;
        private OpenTypeFontEngine _engine;
        private TextShaper _shaper;

        [GlobalSetup]  // Runs once before all benchmarks
        public void Setup()
        {
            var fontsPath = Path.Combine(AppContext.BaseDirectory, "Fonts");
            _engine = new OpenTypeFontEngine(x =>
            {
                x.FontDirectories.Add(fontsPath);
                x.SearchSystemDirectories = false;
            });

            if (!Directory.Exists(fontsPath))
            {
                throw new DirectoryNotFoundException($"Fonts directory not found: {fontsPath}");
            }

            var fontFolders = new List<string> { fontsPath };
            _roboto = _engine.LoadFont("Roboto", FontSubFamily.Regular);
            _shaper = new TextShaper(_engine, _roboto);
        }

        [Benchmark]
        public ShapedText Shape_ShortText()
        {
            return _shaper.Shape("Hello");
        }

        [Benchmark]
        public ShapedText Shape_WithLigatures()
        {
            return _shaper.Shape("office fit");
        }

        [Benchmark]
        public ShapedText Shape_LongText()
        {
            return _shaper.Shape("The quick brown fox jumps over the lazy dog. Office 2024.");
        }

        [Benchmark]
        public ShapedText Shape_LotsOfLigatures()
        {
            return _shaper.Shape("office efficientaffinityuffice");
        }
    }
}
