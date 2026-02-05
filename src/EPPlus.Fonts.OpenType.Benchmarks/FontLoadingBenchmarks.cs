/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2025         EPPlus Software AB           Initial implementation
 *************************************************************************************************/
using BenchmarkDotNet.Attributes;
using EPPlus.Fonts.OpenType;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.Benchmarks
{
    [MemoryDiagnoser]
    [SimpleJob(warmupCount: 3, iterationCount: 5)]
    public class FontLoadingBenchmarks
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
        public OpenTypeFont Load_Roboto_Regular_ColdCache()
        {
            OpenTypeFonts.ClearFontCache(); // Clear INNE i benchmark
            return OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular);
        }

        [Benchmark]
        public OpenTypeFont Load_Roboto_Regular_WarmCache()
        {
            // Load UTAN att cleara - använder cache från GlobalSetup eller warmup
            return OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular);
        }

        [Benchmark]
        public OpenTypeFont Load_Roboto_Bold_ColdCache()
        {
            OpenTypeFonts.ClearFontCache();
            return OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Bold);
        }

        [Benchmark]
        public OpenTypeFont Load_Roboto_Italic_ColdCache()
        {
            OpenTypeFonts.ClearFontCache();
            return OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Italic);
        }

        [Benchmark]
        public OpenTypeFont Load_Roboto_BoldItalic_ColdCache()
        {
            OpenTypeFonts.ClearFontCache();
            return OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.BoldItalic);
        }
    }
}