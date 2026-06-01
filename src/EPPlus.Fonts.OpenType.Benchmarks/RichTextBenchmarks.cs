/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/23/2026         EPPlus Software AB           Debug NA benchmarks
  05/06/2026         EPPlus Software AB           Use OpenTypeFonts.Configure for font directories
 *************************************************************************************************/
using BenchmarkDotNet.Attributes;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.Benchmarks
{
    [MemoryDiagnoser]
    [SimpleJob(warmupCount: 1, iterationCount: 2)]
    public class DebugNABenchmarks
    {
        private const string LoremIpsum20Para = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla pulvinar interdum imperdiet. Praesent ut auctor urna. Phasellus sollicitudin quam vitae est convallis, eu mattis lorem efficitur. Mauris nulla libero, tincidunt id ipsum non, lobortis tristique mauris. Donec ut enim sed enim fermentum molestie vel quis odio. Morbi a fermentum massa, sit amet ultrices est. Aenean ante mi, fermentum nec rhoncus et, vulputate vel sapien. Donec tempus, leo quis luctus rhoncus, augue odio pharetra libero, ac blandit urna turpis sed diam. Vivamus augue purus, eleifend et justo facilisis, imperdiet rhoncus sem. Quisque accumsan pellentesque elit, eget finibus massa accumsan in.\r\n\r\nFusce eu accumsan enim. Cras pulvinar enim vel tellus lacinia, consectetur euismod tortor consectetur. Praesent tincidunt pretium eros, ac auctor magna luctus sed. Ut porta lectus quam, non ornare mauris lacinia sit amet. Nullam egestas dolor quis magna porttitor, ac iaculis nisi hendrerit. Proin at mollis lacus, in porttitor nunc. Aliquam erat volutpat. Sed vel egestas risus, at aliquam arcu. Vestibulum quis lobortis nulla. Etiam pellentesque auctor nulla, eget tincidunt felis rhoncus id.";

        private TextLayoutEngine _layoutEngine;

        private const double MaxPointWidth = 39d;
        private const float FontSize = 11f;
        private const string FontFamily = "Roboto";

        private List<TextFragment> _fragments10;

        [GlobalSetup]
        public void Setup()
        {
            Console.WriteLine("========================================");
            Console.WriteLine("DEBUG NA BENCHMARKS - GlobalSetup START");
            Console.WriteLine("========================================");

            var fontsPath = Path.Combine(AppContext.BaseDirectory, "Fonts");
            Console.WriteLine(string.Format("Fonts directory: {0}", fontsPath));

            if (!Directory.Exists(fontsPath))
            {
                throw new DirectoryNotFoundException(
                    string.Format("Fonts directory not found: {0}", fontsPath));
            }

            // Configure the global font system to search the benchmark's local Fonts directory
            // exclusively. Must happen before any LoadFont call.
            var fontEngine = new OpenTypeFontEngine(cfg =>
            {
                cfg.Reset();
                cfg.FontDirectories.Add(fontsPath);
                cfg.SearchSystemDirectories = false;
            });


            Console.WriteLine("\nAvailable Roboto fonts:");
            foreach (var file in Directory.GetFiles(fontsPath, "Roboto*.ttf"))
            {
                Console.WriteLine(string.Format("  {0}", Path.GetFileName(file)));
            }

            Console.WriteLine("\nLoading Roboto Regular...");
            var font = OpenTypeFonts.LoadFont(FontFamily, FontSubFamily.Regular);

            Console.WriteLine(string.Format("Loaded: {0} {1} ({2} glyphs)",
                font.FullName, font.SubFamily, font.GlyfTable.Glyphs.Count));

            var shaper = new TextShaper(fontEngine, font);
            _layoutEngine = new TextLayoutEngine(shaper);

            Console.WriteLine("\nPre-warming font cache (Regular, Bold, Italic)...");
            PrewarmFontCache();

            Console.WriteLine("\nPreparing 10 test fragments...");
            _fragments10 = new List<TextFragment>();
            var measurementFont = new MeasurementFont
            {
                FontFamily = FontFamily,
                Size = FontSize,
                Style = MeasurementFontStyles.Regular
            };

            for (int i = 0; i < 10; i++)
            {
                var tf = new TextFragment
                {
                    Text = LoremIpsum20Para,
                };
                tf.RichTextOptions.SetFont(measurementFont);
                _fragments10.Add(tf);
            }

            Console.WriteLine(string.Format("Prepared {0} fragments, each with {1} chars",
                _fragments10.Count, LoremIpsum20Para.Length));

            Console.WriteLine("\n=== TESTING BENCHMARK 1 ===");
            try
            {
                var sw = System.Diagnostics.Stopwatch.StartNew();
                var result = TestBenchmark1();
                sw.Stop();
                Console.WriteLine(string.Format("SUCCESS: {0} lines in {1}ms",
                    result.Count, sw.ElapsedMilliseconds));
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("FAILED: {0}", ex.Message));
                Console.WriteLine(ex.StackTrace);
            }

            Console.WriteLine("\n=== TESTING BENCHMARK 2 ===");
            try
            {
                var sw = System.Diagnostics.Stopwatch.StartNew();
                var result = TestBenchmark2();
                sw.Stop();
                Console.WriteLine(string.Format("SUCCESS: {0} lines in {1}ms",
                    result.Count, sw.ElapsedMilliseconds));
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("FAILED: {0}", ex.Message));
                Console.WriteLine(ex.StackTrace);
            }

            Console.WriteLine("\n========================================");
            Console.WriteLine("GlobalSetup COMPLETE");
            Console.WriteLine("========================================\n");
        }

        private void PrewarmFontCache()
        {
            var warmupFragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "warmup",
                    Font = new OpenTypeFontInfoBase
                    {
                        Family = FontFamily,
                        Size = FontSize,
                        SubFamily = FontSubFamily.Regular
                    }
                },
                new TextFragment
                {
                    Text = "warmup",
                    Font = new OpenTypeFontInfoBase
                    {
                        Family = FontFamily,
                        Size = 12f,
                        SubFamily = FontSubFamily.Bold
                    }
                },
                new TextFragment
                {
                    Text = "warmup",
                    Font = new OpenTypeFontInfoBase
                    {
                        Family = FontFamily,
                        Size = FontSize,
                        SubFamily = FontSubFamily.Italic
                    }
                }
            };

            var result = _layoutEngine.WrapRichText(warmupFragments, MaxPointWidth);
            Console.WriteLine(string.Format("  Cache warmed ({0} lines)", result.Count));
        }

        private List<string> TestBenchmark1()
        {
            Console.WriteLine("  Wrapping 10 paragraphs sequentially...");
            List<string> allLines = new List<string>();
            for (int i = 0; i < 10; i++)
            {
                Console.WriteLine(string.Format("    Paragraph {0}...", i + 1));
                var lines = _layoutEngine.WrapRichText(
                    new List<TextFragment> { _fragments10[i] },
                    MaxPointWidth
                );
                allLines.AddRange(lines);
            }
            return allLines;
        }

        private List<string> TestBenchmark2()
        {
            Console.WriteLine("  Creating 5 fragments with mixed fonts...");
            var text = LoremIpsum20Para;
            int chunkSize = text.Length / 5;

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = text.Substring(0, chunkSize),
                    Font = new OpenTypeFontInfoBase { Family = FontFamily, Size = FontSize }
                },
                new TextFragment
                {
                    Text = text.Substring(chunkSize, chunkSize),
                    Font = new OpenTypeFontInfoBase { Family = FontFamily, Size = 12f, SubFamily = FontSubFamily.Bold }
                },
                new TextFragment
                {
                    Text = text.Substring(chunkSize * 2, chunkSize),
                    Font = new OpenTypeFontInfoBase { Family = FontFamily, Size = FontSize, SubFamily = FontSubFamily.Italic }
                },
                new TextFragment
                {
                    Text = text.Substring(chunkSize * 3, chunkSize),
                    Font = new OpenTypeFontInfoBase { Family = FontFamily, Size = FontSize }
                },
                new TextFragment
                {
                    Text = text.Substring(chunkSize * 4),
                    Font = new OpenTypeFontInfoBase { Family = FontFamily, Size = 10f }
                }
            };

            Console.WriteLine(string.Format("  Created {0} fragments", fragments.Count));
            Console.WriteLine("  Wrapping...");
            return _layoutEngine.WrapRichText(fragments, MaxPointWidth);
        }

        [Benchmark]
        public List<string> Wrap_10Paragraphs_RichText()
        {
            List<string> allLines = new List<string>();
            for (int i = 0; i < 10; i++)
            {
                var lines = _layoutEngine.WrapRichText(
                    new List<TextFragment> { _fragments10[i] },
                    MaxPointWidth
                );
                allLines.AddRange(lines);
            }
            return allLines;
        }

        [Benchmark]
        public List<string> WrapRichText_MixedFonts_LongText()
        {
            var text = LoremIpsum20Para;
            int chunkSize = text.Length / 5;

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = text.Substring(0, chunkSize),
                    Font = new OpenTypeFontInfoBase { Family = FontFamily, Size = FontSize }
                },
                new TextFragment
                {
                    Text = text.Substring(chunkSize, chunkSize),
                    Font = new OpenTypeFontInfoBase { Family = FontFamily, Size = 12f, SubFamily = FontSubFamily.Bold }
                },
                new TextFragment
                {
                    Text = text.Substring(chunkSize * 2, chunkSize),
                    Font = new OpenTypeFontInfoBase { Family = FontFamily, Size = FontSize, SubFamily = FontSubFamily.Italic }
                },
                new TextFragment
                {
                    Text = text.Substring(chunkSize * 3, chunkSize),
                    Font = new OpenTypeFontInfoBase { Family = FontFamily, Size = FontSize }
                },
                new TextFragment
                {
                    Text = text.Substring(chunkSize * 4),
                    Font = new OpenTypeFontInfoBase { Family = FontFamily, Size = 10f }
                }
            };

            return _layoutEngine.WrapRichText(fragments, MaxPointWidth);
        }
    }
}