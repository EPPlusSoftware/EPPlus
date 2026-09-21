using BenchmarkDotNet.Attributes;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.Benchmarks
{
    [MemoryDiagnoser]
    [SimpleJob(warmupCount: 3, iterationCount: 5)]
    public class FontLoadingBenchmarks
    {
        private List<string> _fontFolders;

        // Kept across iterations: its font cache is warm after the first load.
        private OpenTypeFontEngine _warmEngine;

        // Rebuilt before every cold iteration by IterationSetup.
        private OpenTypeFontEngine _coldEngine;

        [GlobalSetup]
        public void Setup()
        {
            var fontsPath = Path.Combine(System.AppContext.BaseDirectory, "Fonts");

            if (!Directory.Exists(fontsPath))
            {
                throw new DirectoryNotFoundException($"Fonts directory not found: {fontsPath}");
            }

            _fontFolders = new List<string> { fontsPath };

            _warmEngine = CreateEngine();
            _warmEngine.LoadFont("Roboto", FontSubFamily.Regular);
        }

        [GlobalCleanup]
        public void Cleanup()
        {
            _warmEngine?.Dispose();
            _coldEngine?.Dispose();
        }

        private OpenTypeFontEngine CreateEngine()
        {
            return new OpenTypeFontEngine(cfg =>
            {
                foreach (var folder in _fontFolders)
                {
                    cfg.FontDirectories.Add(folder);
                }
                // The benchmark measures loading the test fonts, not whatever happens to be
                // installed on the machine running it.
                cfg.SearchSystemDirectories = false;
            });
        }

        [IterationSetup(Target = nameof(Load_ColdCache))]
        public void ColdSetup()
        {
            _coldEngine?.Dispose();
            _coldEngine = CreateEngine();
        }

        [Benchmark]
        [Arguments(FontSubFamily.Regular)]
        [Arguments(FontSubFamily.Bold)]
        [Arguments(FontSubFamily.Italic)]
        [Arguments(FontSubFamily.BoldItalic)]
        public OpenTypeFont Load_ColdCache(FontSubFamily subFamily)
        {
            return _coldEngine.LoadFont("Roboto", subFamily);
        }

        [Benchmark]
        public OpenTypeFont Load_Roboto_Regular_WarmCache()
        {
            return _warmEngine.LoadFont("Roboto", FontSubFamily.Regular);
        }

        [Benchmark]
        public OpenTypeFont[] Load_FromCache_AllSubFamilies()
        {
            // Four distinct cache keys in sequence — the pattern a document with all four
            // styles produces. Unlike the single-font warm benchmark, this measures four
            // separate lookups rather than one repeated.
            return new[]
            {
                _warmEngine.LoadFont("Roboto", FontSubFamily.Regular),
                _warmEngine.LoadFont("Roboto", FontSubFamily.Bold),
                _warmEngine.LoadFont("Roboto", FontSubFamily.Italic),
                _warmEngine.LoadFont("Roboto", FontSubFamily.BoldItalic)
            };
        }
    }
}