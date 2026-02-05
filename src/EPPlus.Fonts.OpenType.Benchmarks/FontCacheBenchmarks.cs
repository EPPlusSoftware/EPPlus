using BenchmarkDotNet.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Fonts.OpenType.Benchmarks
{
    /// <summary>
    /// Separate benchmark class to measure cache performance without ClearCache in IterationSetup
    /// </summary>
    [MemoryDiagnoser]
    [SimpleJob(warmupCount: 3, iterationCount: 5)]
    public class FontCacheBenchmarks
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

            // Pre-load font into cache
            OpenTypeFonts.ClearFontCache();
            OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular);
        }

        [Benchmark]
        public OpenTypeFont Load_FromCache_SingleThread()
        {
            // This should be extremely fast - just cache lookup
            return OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular);
        }

        [Benchmark]
        public OpenTypeFont[] Load_FromCache_MultipleFonts()
        {
            // Simulates loading multiple font styles (like for a document)
            return new[]
            {
                OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular),
                OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Bold),
                OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Italic),
                OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.BoldItalic)
            };
        }
    }
}
