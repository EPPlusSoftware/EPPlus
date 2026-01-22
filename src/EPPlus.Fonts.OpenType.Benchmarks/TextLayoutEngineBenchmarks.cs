/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/22/2026         EPPlus Software AB           New TextLayoutEngine benchmarks
 *************************************************************************************************/
using BenchmarkDotNet.Attributes;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Collections.Generic;

namespace EPPlus.Fonts.Benchmarks
{
    /// <summary>
    /// Benchmarks for new TextLayoutEngine text wrapping performance.
    /// Compares with old FontMeasurerTrueType implementation.
    /// </summary>
    [MemoryDiagnoser]
    [SimpleJob(warmupCount: 3, iterationCount: 5)]
    public class TextLayoutEngineBenchmarks
    {
        // 20 paragraphs of 'lorem ipsum' - same as TextMeasurementBenchmarks
        private const string LoremIpsum20Para = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla pulvinar interdum imperdiet. Praesent ut auctor urna. Phasellus sollicitudin quam vitae est convallis, eu mattis lorem efficitur. Mauris nulla libero, tincidunt id ipsum non, lobortis tristique mauris. Donec ut enim sed enim fermentum molestie vel quis odio. Morbi a fermentum massa, sit amet ultrices est. Aenean ante mi, fermentum nec rhoncus et, vulputate vel sapien. Donec tempus, leo quis luctus rhoncus, augue odio pharetra libero, ac blandit urna turpis sed diam. Vivamus augue purus, eleifend et justo facilisis, imperdiet rhoncus sem. Quisque accumsan pellentesque elit, eget finibus massa accumsan in.\r\n\r\nFusce eu accumsan enim. Cras pulvinar enim vel tellus lacinia, consectetur euismod tortor consectetur. Praesent tincidunt pretium eros, ac auctor magna luctus sed. Ut porta lectus quam, non ornare mauris lacinia sit amet. Nullam egestas dolor quis magna porttitor, ac iaculis nisi hendrerit. Proin at mollis lacus, in porttitor nunc. Aliquam erat volutpat. Sed vel egestas risus, at aliquam arcu. Vestibulum quis lobortis nulla. Etiam pellentesque auctor nulla, eget tincidunt felis rhoncus id. Sed metus ante, efficitur id dui eu, fermentum mollis odio. Phasellus ullamcorper iaculis augue vel consequat. Etiam fringilla euismod interdum. Ut molestie massa id fringilla lobortis. Vestibulum malesuada, ante vel mattis ultrices, sem ante molestie augue, non tristique dui mi non nibh.\r\n\r\nMaecenas dictum, sem eget convallis rhoncus, lacus enim porta neque, in posuere dui ex a sapien. Nam lacus nibh, posuere sed elit eget, condimentum facilisis ligula. Cras consectetur lacus ullamcorper velit aliquet bibendum eget vel nulla. Aenean varius ac erat quis ullamcorper. Donec laoreet arcu a lorem volutpat faucibus. Vivamus vehicula leo ut erat luctus scelerisque. Morbi posuere ex et magna egestas facilisis. Fusce scelerisque volutpat erat bibendum hendrerit. Nam blandit mi ut metus pulvinar, vel tempus lacus euismod. Quisque imperdiet sit amet sapien sed ultricies. Phasellus sodales, ipsum vitae tincidunt facilisis, nulla ligula faucibus felis, eget vehicula ante lacus eu lorem.\r\n\r\nInteger congue diam ac viverra tristique. Curabitur tristique dolor quis quam pretium, et scelerisque quam dictum. Maecenas vitae sodales ligula. Pellentesque maximus diam vel porta convallis. Ut aliquam eros quis porta pellentesque. Fusce in ex ut mi egestas cursus. Aliquam erat volutpat. Cras laoreet condimentum laoreet.\r\n\r\nSed eget facilisis tellus. Morbi viverra odio sed odio placerat mollis. Duis turpis metus, dignissim varius urna quis, viverra dignissim dui. Vivamus viverra at nisi quis convallis. Suspendisse fringilla risus et ante sollicitudin, sed eleifend sem placerat. Proin pretium blandit arcu, eget rhoncus risus hendrerit at. Interdum et malesuada fames ac ante ipsum primis in faucibus. Phasellus vulputate efficitur maximus.\r\n\r\nCras blandit nulla eu nisi auctor tempus. Sed pretium lacus ac magna vestibulum, aliquam faucibus orci luctus. Mauris enim lorem, varius ut ante quis, varius viverra lectus. Fusce blandit nibh vel feugiat efficitur. Donec maximus id justo ac mollis. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae; Nulla placerat lectus et purus dictum, id congue nisi euismod. Maecenas euismod fermentum diam, sit amet gravida magna suscipit a. Quisque consectetur arcu eu nunc sodales scelerisque. Nulla non tincidunt nulla. Pellentesque ut tortor vel enim convallis malesuada.\r\n\r\nAliquam ultricies bibendum ultrices. Mauris rutrum ac nisl vel luctus. Donec quis nibh vitae orci ultricies gravida. Aliquam vitae velit porttitor lorem bibendum fringilla volutpat a eros. Curabitur at commodo tortor. Etiam ultricies, neque et iaculis euismod, diam ligula luctus mi, vitae lobortis felis lorem eu nulla. Sed a semper ex. Interdum et malesuada fames ac ante ipsum primis in faucibus. Nulla mauris elit, pulvinar ac tortor et, luctus hendrerit nisl. In egestas auctor urna vitae laoreet. Praesent bibendum egestas convallis. Proin non suscipit tellus.\r\n\r\nNullam at nibh in urna laoreet sodales non vel tellus. Donec in enim dui. Phasellus quis quam tincidunt, pellentesque lorem ac, scelerisque neque. Integer nec tempus urna. Donec elit massa, eleifend eu sapien sit amet, mollis pellentesque est. Nullam tristique tellus iaculis arcu consectetur pretium. Sed venenatis convallis scelerisque. Suspendisse varius urna sit amet purus accumsan, id ultricies erat efficitur. Cras non ipsum eget nulla efficitur commodo sit amet non lacus. Proin viverra enim sit amet enim tempus ullamcorper. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Duis ac massa interdum, gravida ex egestas, finibus purus. Nunc consectetur commodo lacus, ac convallis quam lobortis eu. Sed convallis tempor commodo. Nulla sed convallis mauris.\r\n\r\nDonec venenatis nisi est, ac ullamcorper mi pretium quis. Donec vitae eros at ipsum interdum scelerisque nec vitae nisi. Sed vestibulum erat ac bibendum dapibus. Morbi nec elit id quam tristique cursus id sed sem. Praesent non ante enim. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Praesent non mauris dui. Aliquam rhoncus mattis ante sed venenatis. Vivamus vehicula sed sapien sed dictum. In aliquet, urna efficitur tincidunt lobortis, nibh justo tristique purus, sed volutpat risus magna et libero.\r\n\r\nSuspendisse lectus justo, varius eget arcu et, semper laoreet erat. Quisque eget lacus ornare, pellentesque erat sit amet, vulputate felis. Duis luctus, massa a pellentesque mollis, massa elit convallis mi, vel bibendum ex ex eu purus. Suspendisse vel fermentum urna, ac commodo enim. Mauris tincidunt cursus elit, a volutpat libero commodo et. Etiam dapibus libero venenatis tellus lobortis, vel lacinia elit faucibus. Maecenas semper sed quam quis finibus. Integer efficitur, libero imperdiet sollicitudin commodo, elit arcu vulputate est, eget finibus mi urna sit amet magna. Cras ullamcorper consequat ornare. Fusce convallis nunc vel risus cursus, at maximus ligula cursus. Pellentesque vulputate risus libero, eget cursus nibh sodales sed. Donec accumsan sem et massa semper, id dignissim velit vehicula.\r\n\r\nCras cursus ipsum ac erat vehicula, nec iaculis purus dictum. Quisque lacinia elit vitae leo dictum, vel dignissim velit dapibus. Aenean sem nisi, faucibus interdum justo eu, euismod porttitor ex. Morbi et lectus lectus. Duis neque felis, suscipit at scelerisque eu, scelerisque id orci. Curabitur et placerat ipsum. Proin gravida sapien nisl, et varius ipsum mollis nec. Quisque dignissim consectetur feugiat. Aenean eros purus, laoreet interdum rutrum at, aliquet sit amet lectus. Donec gravida lorem ut tincidunt laoreet. Donec consequat viverra ligula, in accumsan mi bibendum scelerisque. Quisque ac risus justo. Morbi magna arcu, egestas nec luctus commodo, cursus eget nunc. Vivamus euismod lorem ex, et maximus felis hendrerit eget. Nullam ullamcorper euismod ligula, et iaculis ligula ultricies a. Fusce aliquam, enim vel fermentum ultrices, elit quam semper erat, vitae semper velit augue non magna.\r\n\r\nQuisque maximus semper arcu, id pellentesque est tempus a. Phasellus lacus elit, auctor sit amet lacinia a, dapibus vitae velit. Phasellus ut pharetra justo, ut ultricies erat. Sed molestie sapien vel interdum lobortis. Nulla facilisi. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae; Nulla nec mauris quis nisi vulputate gravida quis nec velit.\r\n\r\nNam et congue ipsum. Nulla vel elit non dolor mollis aliquet vel at magna. Pellentesque nec facilisis elit. In vulputate quis sem porta suscipit. Nullam sed ex ornare nibh suscipit mattis quis non lacus. Mauris vel ex urna. Vivamus ultricies sapien sit amet sapien vehicula gravida. Donec feugiat volutpat quam. Vestibulum auctor dictum nisl, id hendrerit metus ullamcorper sed. Nulla maximus lacus vel mollis maximus. Nulla laoreet placerat quam eu viverra. Etiam feugiat accumsan nisl a condimentum. Sed ultricies ante ante, ac auctor ligula gravida nec. Praesent a neque dignissim, sagittis felis sit amet, condimentum turpis.\r\n\r\nFusce at leo vel est blandit malesuada. Pellentesque et neque non metus pellentesque imperdiet. Praesent pellentesque lacinia lorem, et tristique tellus efficitur id. Suspendisse aliquet ultricies justo vitae interdum. Cras tristique viverra quam, eget gravida mi fermentum imperdiet. Sed imperdiet vitae purus ut volutpat. Nulla lacinia elit in fermentum consectetur. Phasellus commodo ut nisl sit amet sagittis. Duis ac ornare orci.\r\n\r\nVivamus vel enim posuere, pharetra ex vel, elementum est. Vestibulum commodo luctus metus eget maximus. Suspendisse a nulla a odio eleifend faucibus. Suspendisse semper lacus non porttitor aliquet. Cras ac scelerisque magna, et pulvinar justo. Integer cursus pulvinar fringilla. Mauris imperdiet nibh sit amet tempor laoreet. Morbi tincidunt tortor ex, sit amet maximus purus tristique quis. Quisque sed hendrerit velit. Mauris mattis nibh ut eros luctus, eget mattis massa auctor. Phasellus eu neque at augue gravida sagittis nec non tortor. Etiam porttitor sem sodales mi ullamcorper gravida.\r\n\r\nIn in dictum orci. In vitae vestibulum quam. Cras augue eros, tincidunt ac elit posuere, sollicitudin efficitur lectus. Praesent quis sodales nisl. Proin sit amet molestie est. In commodo mauris vel mauris efficitur, nec mollis mauris sagittis. Cras ligula nibh, egestas sit amet eros in, lacinia tristique magna. Cras risus libero, lacinia eget libero vitae, maximus aliquet nibh. Mauris id sodales purus, vitae dictum lectus. Cras consectetur ligula velit, tempus pulvinar lacus porttitor vitae. Phasellus eget tellus ipsum.\r\n\r\nDonec interdum laoreet elit non vestibulum. Cras sed urna ullamcorper, aliquam erat eget, porta orci. Vestibulum eget congue nulla. Sed sem tortor, euismod at rutrum id, sagittis a nunc. Duis in nibh facilisis, dignissim purus ut, hendrerit magna. Sed semper ligula id massa elementum, non malesuada velit egestas. Nullam dictum, mi nec euismod sagittis, ligula leo ullamcorper dolor, quis faucibus odio metus eget magna. Ut gravida metus non metus bibendum bibendum. In sagittis eleifend aliquet.\r\n\r\nInterdum et malesuada fames ac ante ipsum primis in faucibus. Nam mollis sagittis felis, in faucibus tortor pretium vel. Nam nec enim metus. Donec in augue arcu. Proin non lobortis purus, sit amet lacinia elit. Suspendisse quis eros condimentum, blandit justo sit amet, lobortis nisl. Suspendisse maximus massa sed urna tempor ornare. Nunc malesuada purus odio, eu luctus lectus auctor nec. Morbi auctor pellentesque auctor. Sed ullamcorper, ex vitae aliquam vulputate, est diam feugiat mi, id porttitor lectus orci ac leo.\r\n\r\nDonec sit amet velit pulvinar, venenatis turpis ut, interdum ligula. Interdum et malesuada fames ac ante ipsum primis in faucibus. Vestibulum eu lacus urna. Maecenas sem nulla, accumsan eu ultricies sed, tempor vel magna. Cras aliquet sollicitudin sapien ac pulvinar. Praesent ac sodales mi. Integer vitae mauris massa. Maecenas iaculis orci et faucibus interdum.\r\n\r\nNunc nec maximus felis, sed finibus quam. Pellentesque felis massa, vestibulum in tellus vitae, congue tincidunt justo. Nunc vitae enim malesuada, bibendum ante nec, varius tellus. Praesent vitae nisi id quam auctor lacinia at non quam. Nam nec ligula sit amet felis auctor sagittis. Nunc in risus eu urna varius laoreet quis sit amet felis. Morbi varius tempor orci, eu vestibulum nunc vestibulum ac. Nunc vehicula velit eleifend consequat porta. Suspendisse maximus dapibus orci, in vulputate massa pretium ac. Quisque malesuada aliquet aliquet.";

        private TextLayoutEngine _layoutEngine;
        private FontMeasurerTrueType _oldMeasurer;

        private const double MaxPixelWidth = 52d;
        private const double MaxPointWidth = 39d; // 52 pixels ≈ 39 points (at 96 DPI)
        private const float FontSize = 11f;
        private const string FontFamily = "Roboto";

        private List<TextFragment> _fragments100;
        private List<string> _texts100;
        private List<MeasurementFont> _fonts100;

        [GlobalSetup]
        public void Setup()
        {
            // Setup old measurer
            _oldMeasurer = new FontMeasurerTrueType();
            _oldMeasurer.SetFont(FontSize, FontFamily);

            // Setup new layout engine
            var font = OpenTypeFonts.GetFontData(null, FontFamily, FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);
            _layoutEngine = new TextLayoutEngine(shaper);

            // Prepare 100 copies of the long text
            _fragments100 = new List<TextFragment>();
            _texts100 = new List<string>();
            _fonts100 = new List<MeasurementFont>();

            var measurementFont = new MeasurementFont
            {
                FontFamily = FontFamily,
                Size = FontSize,
                Style = MeasurementFontStyles.Regular
            };

            for (int i = 0; i < 100; i++)
            {
                _texts100.Add(LoremIpsum20Para);
                _fonts100.Add(measurementFont);
                _fragments100.Add(new TextFragment
                {
                    Text = LoremIpsum20Para,
                    Font = measurementFont
                });
            }
        }

        #region Old Implementation Benchmarks (Baseline)

        [Benchmark(Baseline = true)]
        public List<string> Old_Wrap_SingleParagraph()
        {
            return _oldMeasurer.MeasureAndWrapText(LoremIpsum20Para, MaxPixelWidth);
        }

        [Benchmark]
        public List<string> Old_Wrap_100Paragraphs_Sequential()
        {
            List<string> wrapped = new List<string>();
            foreach (string text in _texts100)
            {
                wrapped = _oldMeasurer.MeasureAndWrapText(text, MaxPixelWidth);
            }
            return wrapped;
        }

        [Benchmark]
        public List<string> Old_Wrap_100Paragraphs_MultipleFragments()
        {
            return _oldMeasurer.WrapMultipleTextFragments(_texts100, _fonts100, MaxPixelWidth);
        }

        [Benchmark]
        public List<string> Old_Wrap_ShortText()
        {
            var shortText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit.";
            return _oldMeasurer.MeasureAndWrapText(shortText, MaxPixelWidth);
        }

        [Benchmark]
        public List<string> Old_Wrap_WideColumn()
        {
            return _oldMeasurer.MeasureAndWrapText(LoremIpsum20Para, 200d);
        }

        [Benchmark]
        public List<string> Old_Wrap_NarrowColumn()
        {
            return _oldMeasurer.MeasureAndWrapText(LoremIpsum20Para, 30d);
        }

        #endregion

        #region New Implementation Benchmarks

        [Benchmark]
        public List<string> New_Wrap_SingleParagraph()
        {
            return _layoutEngine.WrapText(LoremIpsum20Para, FontSize, MaxPointWidth);
        }

        [Benchmark]
        public List<string> New_Wrap_100Paragraphs_Sequential()
        {
            List<string> wrapped = new List<string>();
            foreach (string text in _texts100)
            {
                wrapped = _layoutEngine.WrapText(text, FontSize, MaxPointWidth);
            }
            return wrapped;
        }

        [Benchmark]
        public double[] OnlyExtractWidths()
        {
            var font = OpenTypeFonts.GetFontData(null, FontFamily, FontSubFamily.Regular, true);
            var shaper = new TextShaper(font);
            return shaper.ExtractCharWidths(LoremIpsum20Para, FontSize, ShapingOptions.Default);
        }

        [Benchmark]
        public List<string> New_Wrap_100Paragraphs_RichText()
        {
            // Note: This wraps each text individually, not as one concatenated text
            // (matching old behavior more closely than wrapping all as single rich text)
            List<string> allLines = new List<string>();
            foreach (var fragment in _fragments100)
            {
                var lines = _layoutEngine.WrapRichText(new List<TextFragment> { fragment }, MaxPointWidth);
                allLines.AddRange(lines);
            }
            return allLines;
        }

        [Benchmark]
        public List<string> New_Wrap_ShortText()
        {
            var shortText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit.";
            return _layoutEngine.WrapText(shortText, FontSize, MaxPointWidth);
        }

        [Benchmark]
        public List<string> New_Wrap_WideColumn()
        {
            return _layoutEngine.WrapText(LoremIpsum20Para, FontSize, 150d); // ~200 pixels in points
        }

        [Benchmark]
        public List<string> New_Wrap_NarrowColumn()
        {
            return _layoutEngine.WrapText(LoremIpsum20Para, FontSize, 22.5d); // ~30 pixels in points
        }

        #endregion

        #region Rich Text Specific Benchmarks

        [Benchmark]
        public List<string> New_WrapRichText_MixedFonts_ShortText()
        {
            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "Lorem ipsum ",
                    Font = new MeasurementFont { FontFamily = FontFamily, Size = FontSize }
                },
                new TextFragment
                {
                    Text = "dolor sit amet, ",
                    Font = new MeasurementFont { FontFamily = FontFamily, Size = 12f, Style = MeasurementFontStyles.Bold }
                },
                new TextFragment
                {
                    Text = "consectetur adipiscing elit.",
                    Font = new MeasurementFont { FontFamily = FontFamily, Size = FontSize }
                }
            };

            return _layoutEngine.WrapRichText(fragments, MaxPointWidth);
        }

        [Benchmark]
        public List<string> New_WrapRichText_MixedFonts_LongText()
        {
            // Split lorem ipsum into 5 fragments with different formatting
            var text = LoremIpsum20Para;
            int chunkSize = text.Length / 5;

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = text.Substring(0, chunkSize),
                    Font = new MeasurementFont { FontFamily = FontFamily, Size = FontSize }
                },
                new TextFragment
                {
                    Text = text.Substring(chunkSize, chunkSize),
                    Font = new MeasurementFont { FontFamily = FontFamily, Size = 12f, Style = MeasurementFontStyles.Bold }
                },
                new TextFragment
                {
                    Text = text.Substring(chunkSize * 2, chunkSize),
                    Font = new MeasurementFont { FontFamily = FontFamily, Size = FontSize, Style = MeasurementFontStyles.Italic }
                },
                new TextFragment
                {
                    Text = text.Substring(chunkSize * 3, chunkSize),
                    Font = new MeasurementFont { FontFamily = FontFamily, Size = FontSize }
                },
                new TextFragment
                {
                    Text = text.Substring(chunkSize * 4),
                    Font = new MeasurementFont { FontFamily = FontFamily, Size = 10f }
                }
            };

            return _layoutEngine.WrapRichText(fragments, MaxPointWidth);
        }

        #endregion
    }
}