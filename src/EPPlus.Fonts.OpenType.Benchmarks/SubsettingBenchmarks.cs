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
    public class SubsettingBenchmarks
    {
        private OpenTypeFont _roboto;
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
            _roboto = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular);
        }

        [Benchmark]
        public OpenTypeFont Subset_SmallText_ABC()
        {
            return _roboto.CreateSubset("abc");
        }

        [Benchmark]
        public OpenTypeFont Subset_SmallText_WithLigatures()
        {
            return _roboto.CreateSubset("office fit");
        }

        [Benchmark]
        public OpenTypeFont Subset_MediumText_Sentence()
        {
            return _roboto.CreateSubset("The quick brown fox jumps over the lazy dog");
        }

        [Benchmark]
        public OpenTypeFont Subset_LargeText_Paragraph()
        {
            return _roboto.CreateSubset(
                "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris."
            );
        }

        [Benchmark]
        public OpenTypeFont Subset_Numbers_And_Symbols()
        {
            return _roboto.CreateSubset("0123456789 +-*/=()[]{},.;:!?@#$%^&");
        }

        [Benchmark]
        public OpenTypeFont Subset_Swedish_Characters()
        {
            return _roboto.CreateSubset("åäöÅÄÖ Sverige Stockholm");
        }

        [Benchmark]
        public OpenTypeFont Subset_MixedContent_Realistic()
        {
            // Simulates a realistic spreadsheet with headers, numbers, and text
            return _roboto.CreateSubset(
                "Product Name Price Quantity Total " +
                "Office Supplies $123.45 100 $12,345.00 " +
                "Furniture & Equipment €987.65 50 €49,382.50 " +
                "Q1 2024 Revenue Summary"
            );
        }

        [Benchmark]
        public OpenTypeFont Subset_AllAscii()
        {
            // All printable ASCII characters (32-126)
            var ascii = "";
            for (int i = 32; i <= 126; i++)
            {
                ascii += (char)i;
            }
            return _roboto.CreateSubset(ascii);
        }

        [Benchmark]
        public OpenTypeFont Subset_RepeatedCharacters()
        {
            // Tests deduplication efficiency
            return _roboto.CreateSubset("aaaabbbbccccddddeeeeffffgggg");
        }

        [Benchmark]
        public OpenTypeFont Subset_FullAlphabet_LowerAndUpper()
        {
            return _roboto.CreateSubset("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ");
        }
    }
}