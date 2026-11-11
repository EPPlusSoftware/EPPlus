/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Export.Pdf.Helpers;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfIconRenderer
    {
        internal static void Render(string iconName, PdfContentStream cs)
        {
            switch (iconName)
            {
                // arrows
                case "RedDownArrow": Arrow(cs, 0, "#d86344", "#9e381c"); break;
                case "YellowSideArrow": Arrow(cs, -90, "#eac282", "#a4802b"); break;
                case "GreenUpArrow": Arrow(cs, 180, "#76a797", "#3f7865"); break;
                case "GrayDownArrow": Arrow(cs, 0, "#808080", "#646262"); break;
                case "GraySideArrow": Arrow(cs, -90, "#808080", "#646262"); break;
                case "GrayUpArrow": Arrow(cs, 180, "#808080", "#646262"); break;
                case "YellowDownInclineArrow": Arrow(cs, -45, "#eac282", "#a4802b"); break;
                case "YellowUpInclineArrow": Arrow(cs, -135, "#eac282", "#a4802b"); break;
                case "GrayDownInclineArrow": Arrow(cs, -45, "#808080", "#646262"); break;
                case "GrayUpInclineArrow": Arrow(cs, -135, "#808080", "#646262"); break;

                // circles
                case "RedCircle":
                case "RedCircleWithBorder": Circle(cs, "#d65532", "#ac563e"); break;
                case "YellowCircle": Circle(cs, "#eac282", "#a88636"); break;
                case "GreenCircle": Circle(cs, "#68a490", "#387360"); break;
                case "BlackCircle":
                case "BlackCircleWithBorder": Circle(cs, "#505050", "#33312f"); break;
                case "GrayCircle": Circle(cs, "#b1b1b1", "#74716e"); break;
                case "PinkCircle": Circle(cs, "#edb9ab", "#b18478"); break;

                // traffic lights
                case "RedTrafficLight": TrafficLight(cs, "#d86344"); break;
                case "YellowTrafficLight": TrafficLight(cs, "#eac282"); break;
                case "GreenTrafficLight": TrafficLight(cs, "#68a490"); break;

                // triangles / dash / diamond
                case "GreenUpTriangle": SmallTriangle(cs, 0, "#76a797", "#326f5b"); break;
                case "RedDownTriangle": SmallTriangle(cs, 180, "#d65532", "#ac563e"); break;
                case "YellowDash": YellowDash(cs); break;
                case "RedDiamond": RedDiamond(cs); break;
                case "YellowTriangle": YellowTriangle(cs); break;

                // flags
                case "RedFlag": Flag(cs, "#d65532", "#ac563e"); break;
                case "YellowFlag": Flag(cs, "#eac282", "#a88636"); break;
                case "GreenFlag": Flag(cs, "#76a797", "#326f5b"); break;

                // standalone symbols
                case "RedCross": Cross(cs, "#d76244", "#a23c20"); break;
                case "YellowExclamation": Exclamation(cs, "#edc87e", "#a78433"); break;
                case "GreenCheck": Checkmark(cs, "#71a392", "#689485"); break;

                // circle + white symbol
                case "RedCrossSymbol":
                    Circle(cs, "#d65532", "#ac563e");
                    cs.AddCommand("q"); PushShrinkAndCenter(cs); Cross(cs, "#fefdfd", "#a23c20"); cs.AddCommand("Q");
                    break;
                case "YellowExclamationSymbol":
                    Circle(cs, "#eac282", "#a88636");
                    cs.AddCommand("q"); PushShrinkAndCenter(cs); Exclamation(cs, "#fefdfd", "a78433"); cs.AddCommand("Q");
                    break;
                case "GreenCheckSymbol":
                    Circle(cs, "#68a490", "#387360");
                    cs.AddCommand("q"); PushShrinkAndCenter(cs); Checkmark(cs, "#fefdfd", "#689485"); cs.AddCommand("Q");
                    break;

                // stars
                case "GoldStar": Star(cs, "#eac282"); break;
                case "SilverStar": Star(cs, "#ffffff"); break;
                case "HalfGoldStar": HalfGoldStar(cs); break;

                // signal meters
                case "SignalMeterWithNoFilledBars": SignalMeter(cs, 0); break;
                case "SignalMeterWithOneFilledBar": SignalMeter(cs, 1); break;
                case "SignalMeterWithTwoFilledBars": SignalMeter(cs, 2); break;
                case "SignalMeterWithThreeFilledBars": SignalMeter(cs, 3); break;
                case "SignalMeterWithFourFilledBars": SignalMeter(cs, 4); break;

                // quarters
                case "WhiteCircle": Quarters(cs, true, true, true, true); break;
                case "CircleWithThreeWhiteQuarters": Quarters(cs, true, true, true, false); break;
                case "CircleWithTwoWhiteQuarters": Quarters(cs, true, false, true, false); break;
                case "CircleWithOneWhiteQuarter": Quarters(cs, false, false, true, false); break;

                // boxes
                case "ZeroFilledBoxes": FilledBoxes(cs, 0); break;
                case "OneFilledBox": FilledBoxes(cs, 1); break;
                case "TwoFilledBoxes": FilledBoxes(cs, 2); break;
                case "ThreeFilledBoxes": FilledBoxes(cs, 3); break;
                case "FourFilledBoxes": FilledBoxes(cs, 4); break;

                case "NoIcon":
                default:
                    break;   // nothing to draw
            }
        }

        private static void PushRotation(PdfContentStream cs, double svgRotation)
        {
            if (svgRotation == 0d) return;
            double rad = -svgRotation * System.Math.PI / 180d;   // baked y-flip reverses rotation sense
            double cos = System.Math.Cos(rad), sin = System.Math.Sin(rad);
            cs.AddCommand($"{cos.ToPdfStringF4()} {sin.ToPdfStringF4()} {(-sin).ToPdfStringF4()} {cos.ToPdfStringF4()} " +
                          $"{(50d * (1d - cos + sin)).ToPdfStringF4()} {(50d * (1d - sin - cos)).ToPdfStringF4()} cm");
        }

        private static void PushShrinkAndCenter(PdfContentStream cs)
        {
            cs.AddCommand("0.6 0 0 0.6 20 20 cm");   // SVG translate(20,20) scale(0.6), flip-invariant
        }

        private static void CirclePath43(PdfContentStream cs)
        {
            cs.AddCommand("93 50 m");
            cs.AddCommand("93 73.748 73.748 93 50 93 c");
            cs.AddCommand("26.252 93 7 73.748 7 50 c");
            cs.AddCommand("7 26.252 26.252 7 50 7 c");
            cs.AddCommand("73.748 7 93 26.252 93 50 c");
            cs.AddCommand("h");
        }

        private static void YellowTriangle(PdfContentStream cs)
        {
            cs.AddCommand(PdfColor.SetColorFromHex("#eac282").ToFillCommand());
            cs.AddCommand(PdfColor.SetColorFromHex("#a78433").ToStrokeCommand());
            cs.AddCommand("2.5 w");
            cs.AddCommand("88.56 6.035 m");     // bottom edge
            cs.AddCommand("7.315 6.235 l");
            cs.AddCommand("2.953 13.833 l");    // left corner bevel
            cs.AddCommand("44.713 85.753 l");   // apex (bevelled)
            cs.AddCommand("51.04 85.738 l");
            cs.AddCommand("92.7 13.165 l");     // right corner bevel
            cs.AddCommand("h");
            cs.AddCommand("B");
        }

        private static void Flag(PdfContentStream cs, string fillHex, string strokeHex)
        {
            cs.AddCommand(PdfColor.SetColorFromHex("#808080").ToFillCommand());   // pole
            cs.AddCommand("26.072 2.729 6.646 84.789 re");
            cs.AddCommand("f");

            cs.AddCommand(PdfColor.SetColorFromHex(fillHex).ToFillCommand());      // pennant
            cs.AddCommand(PdfColor.SetColorFromHex(strokeHex).ToStrokeCommand());
            cs.AddCommand("1.3 w");
            cs.AddCommand("85.988 63.238 m");
            cs.AddCommand("35.46 39.499 l");
            cs.AddCommand("35.149 86.542 l");
            cs.AddCommand("h");
            cs.AddCommand("B");
        }

        private static void FilledBoxes(PdfContentStream cs, int numActive)
        {
            cs.AddCommand(PdfColor.SetColorFromHex("#757575").ToFillCommand());   // center dot
            cs.AddCommand("46.176 45.436 7.834 7.8854 re");
            cs.AddCommand("f");

            var boxes = new (double x, double y)[]
            {
                (0.68,   0.69),     // bottom-left
                (50.562, 0.689),    // bottom-right
                (0.875,  50.233),   // top-left
                (50.563, 50.234),   // top-right
            };
            for (int i = 0; i < 4; i++)
            {
                bool active = i < numActive;
                cs.AddCommand(PdfColor.SetColorFromHex(active ? "#4d82b8" : "#b3b3b3").ToFillCommand());
                cs.AddCommand(PdfColor.SetColorFromHex(active ? "#335f8c" : "#757575").ToStrokeCommand());
                cs.AddCommand("2.2 w");
                cs.AddCommand($"{boxes[i].x.ToPdfString()} {boxes[i].y.ToPdfString()} 48.757 49.077 re");
                cs.AddCommand("B");
            }
        }

        private static void Quarters(PdfContentStream cs, bool q0, bool q1, bool q2, bool q3)
        {
            cs.AddCommand(PdfColor.SetColorFromHex("#505050").ToFillCommand());   // dark base disc
            CirclePath43(cs);
            cs.AddCommand("f");

            cs.AddCommand(PdfColor.SetColorFromHex("#ffffff").ToFillCommand());   // shown white quarters
            if (q0) { cs.AddCommand("6 51 m"); cs.AddCommand("6 27.252 27.252 6 51 6 c"); cs.AddCommand("51 51 l"); cs.AddCommand("h"); cs.AddCommand("f"); }
            if (q1) { cs.AddCommand("49 51 m"); cs.AddCommand("49 6 l"); cs.AddCommand("72.748 6 94 27.252 94 51 c"); cs.AddCommand("h"); cs.AddCommand("f"); }
            if (q2) { cs.AddCommand("51 94 m"); cs.AddCommand("27.252 94 6 72.748 6 49 c"); cs.AddCommand("51 49 l"); cs.AddCommand("h"); cs.AddCommand("f"); }
            if (q3) { cs.AddCommand("49 94 m"); cs.AddCommand("49 49 l"); cs.AddCommand("94 49 l"); cs.AddCommand("94 72.748 72.748 94 49 94 c"); cs.AddCommand("h"); cs.AddCommand("f"); }

            cs.AddCommand(PdfColor.SetColorFromHex("#33312f").ToStrokeCommand());  // ring outline
            cs.AddCommand("6 w");
            CirclePath43(cs);
            cs.AddCommand("S");
        }

        private static void SignalMeter(PdfContentStream cs, int numActive)
        {
            var bars = new (double x, double y, double w, double h)[]
            {
                (0.602,  0.601, 19.523, 41.049),
                (21.207, 0.727, 19.273, 60.664),
                (41.789, 0.82,  19.07,  79.623),
                (62.357, 0.904, 18.895, 98.18),
            };
            for (int i = 0; i < 4; i++)
            {
                bool active = i < numActive;
                cs.AddCommand(PdfColor.SetColorFromHex(active ? "#4d82b8" : "#b3b3b3").ToFillCommand());
                cs.AddCommand(PdfColor.SetColorFromHex(active ? "#335f8c" : "#757575").ToStrokeCommand());
                cs.AddCommand("3.2 w");
                cs.AddCommand($"{bars[i].x.ToPdfString()} {bars[i].y.ToPdfString()} {bars[i].w.ToPdfString()} {bars[i].h.ToPdfString()} re");
                cs.AddCommand("B");
            }
        }

        private static void HalfGoldStar(PdfContentStream cs)
        {
            cs.AddCommand("q"); cs.AddCommand("0 0 50 100 re W n");       // left half gold
            cs.AddCommand(PdfColor.SetColorFromHex("#eac282").ToFillCommand());
            StarPath(cs); cs.AddCommand("f"); cs.AddCommand("Q");

            cs.AddCommand("q"); cs.AddCommand("50 0 50 100 re W n");      // right half white
            cs.AddCommand(PdfColor.SetColorFromHex("#ffffff").ToFillCommand());
            StarPath(cs); cs.AddCommand("f"); cs.AddCommand("Q");

            cs.AddCommand(PdfColor.SetColorFromHex("#a6812d").ToStrokeCommand());   // single outline
            cs.AddCommand("3.1177 w");
            StarPath(cs); cs.AddCommand("S");
        }

        private static void Star(PdfContentStream cs, string fillHex)
        {
            cs.AddCommand(PdfColor.SetColorFromHex(fillHex).ToFillCommand());
            cs.AddCommand(PdfColor.SetColorFromHex("#a6812d").ToStrokeCommand());
            cs.AddCommand("3.1177 w");
            StarPath(cs);
            cs.AddCommand("B");
        }

        private static void StarPath(PdfContentStream cs)
        {
            cs.AddCommand("49.75 98.381 m");
            cs.AddCommand("47.958 98.381 38.645 64.279 37.195 63.174 c");
            cs.AddCommand("35.745 62.069 1.938 63.321 1.384 61.533 c");
            cs.AddCommand("0.83 59.746 28.881 39.92 29.435 38.133 c");
            cs.AddCommand("29.989 36.346 18.407 3.015 19.857 1.91 c");
            cs.AddCommand("21.307 0.805 47.957 22.654 49.75 22.654 c");
            cs.AddCommand("51.542 22.654 78.191 0.805 79.641 1.91 c");
            cs.AddCommand("81.091 3.015 69.509 36.345 70.063 38.133 c");
            cs.AddCommand("70.617 39.921 98.67 59.746 98.116 61.533 c");
            cs.AddCommand("97.562 63.321 63.755 62.069 62.305 63.174 c");
            cs.AddCommand("60.855 64.279 51.543 98.381 49.75 98.381 c");
            cs.AddCommand("h");
        }

        private static void Checkmark(PdfContentStream cs, string fillHex, string strokeHex)
        {
            cs.AddCommand(PdfColor.SetColorFromHex(fillHex).ToFillCommand());
            cs.AddCommand(PdfColor.SetColorFromHex(strokeHex).ToStrokeCommand());
            cs.AddCommand("3.2 w");
            cs.AddCommand("80.752 98.029 m");
            cs.AddCommand("39.766 34.705 l");
            cs.AddCommand("15.721 58.383 l");
            cs.AddCommand("1.854 44.457 l");
            cs.AddCommand("39.165 2.937 l");
            cs.AddCommand("60.494 28.792 78.719 56.954 98.421 84.142 c");
            cs.AddCommand("h");
            cs.AddCommand("B");
        }

        private static void Exclamation(PdfContentStream cs, string fillHex, string strokeHex)
        {
            cs.AddCommand(PdfColor.SetColorFromHex(fillHex).ToFillCommand());
            cs.AddCommand(PdfColor.SetColorFromHex(strokeHex).ToStrokeCommand());
            cs.AddCommand("1.5863 w");
            cs.AddCommand("40.836 99.207 m");                  // bar (tapered)
            cs.AddCommand("63.558 99.207 l");
            cs.AddCommand("59.59 23.622 l");
            cs.AddCommand("44.798 23.622 l");
            cs.AddCommand("h");
            cs.AddCommand("B");
            cs.AddCommand("62.534 10.774 m");                  // dot (ellipse rx10.086 ry9.738 -> 4 cubics)
            cs.AddCommand("62.534 16.152 58.018 20.512 52.448 20.512 c");
            cs.AddCommand("46.878 20.512 42.362 16.152 42.362 10.774 c");
            cs.AddCommand("42.362 5.396 46.878 1.036 52.448 1.036 c");
            cs.AddCommand("58.018 1.036 62.534 5.396 62.534 10.774 c");
            cs.AddCommand("h");
            cs.AddCommand("B");
        }

        private static void Cross(PdfContentStream cs, string fillHex, string strokeHex)
        {
            cs.AddCommand(PdfColor.SetColorFromHex(fillHex).ToFillCommand());
            cs.AddCommand(PdfColor.SetColorFromHex(strokeHex).ToStrokeCommand());
            cs.AddCommand("3.2963 w");
            cs.AddCommand("20.415 97.669 m");
            cs.AddCommand("2.726 80.821 l");
            cs.AddCommand("33.017 50.186 l");
            cs.AddCommand("2.33 19.137 l");
            cs.AddCommand("20.006 2.289 l");
            cs.AddCommand("50.182 32.824 l");
            cs.AddCommand("79.936 2.728 l");
            cs.AddCommand("97.625 19.564 l");
            cs.AddCommand("67.359 50.187 l");
            cs.AddCommand("97.216 80.396 l");
            cs.AddCommand("79.54 97.244 l");
            cs.AddCommand("50.182 67.55 l");
            cs.AddCommand("h");
            cs.AddCommand("B");
        }

        private static void RedDiamond(PdfContentStream cs)
        {
            cs.AddCommand(PdfColor.SetColorFromHex("#d65532").ToFillCommand());
            cs.AddCommand(PdfColor.SetColorFromHex("#a95139").ToStrokeCommand());
            cs.AddCommand("3 w");
            cs.AddCommand("49.498 92.929 m");                 // square, rotate(45) about origin, flipped
            cs.AddCommand("95.46 46.967 l");
            cs.AddCommand("49.498 1 l");
            cs.AddCommand("3.536 46.967 l");
            cs.AddCommand("h");
            cs.AddCommand("B");
        }

        private static void YellowDash(PdfContentStream cs)
        {
            cs.AddCommand(PdfColor.SetColorFromHex("#edc87e").ToFillCommand());
            cs.AddCommand(PdfColor.SetColorFromHex("#a4802c").ToStrokeCommand());
            cs.AddCommand("3.5893 w");
            cs.AddCommand("1.7947 35.06 96.41 33.145 re");    // y-flipped rect
            cs.AddCommand("B");
        }

        private static void SmallTriangle(PdfContentStream cs, double svgRotation, string fillHex, string strokeHex)
        {
            cs.AddCommand("q");
            PushRotation(cs, svgRotation);
            cs.AddCommand(PdfColor.SetColorFromHex(fillHex).ToFillCommand());
            cs.AddCommand(PdfColor.SetColorFromHex(strokeHex).ToStrokeCommand());
            cs.AddCommand("3.2257 w");
            cs.AddCommand("49.984 68.545 m");
            cs.AddCommand("49.018 68.545 1.188 9.249 1.672 8.65 c");
            cs.AddCommand("2.155 8.051 97.812 8.051 98.295 8.65 c");
            cs.AddCommand("98.778 9.249 50.951 68.545 49.984 68.545 c");
            cs.AddCommand("h");
            cs.AddCommand("B");
            cs.AddCommand("Q");
        }

        private static void TrafficLight(PdfContentStream cs, string fillHex)
        {
            cs.AddCommand("q");
            cs.AddCommand("0 0 100 100 re W n");             // clip the thick overflowing stroke to the box
            cs.AddCommand(PdfColor.SetColorFromHex(fillHex).ToFillCommand());
            cs.AddCommand(PdfColor.SetColorFromHex("#505050").ToStrokeCommand());
            cs.AddCommand("32 w");
            cs.AddCommand("100 50 m");                        // r50 circle -> 4 cubics (k·r≈27.614)
            cs.AddCommand("100 77.614 77.614 100 50 100 c");
            cs.AddCommand("22.386 100 0 77.614 0 50 c");
            cs.AddCommand("0 22.386 22.386 0 50 0 c");
            cs.AddCommand("77.614 0 100 22.386 100 50 c");
            cs.AddCommand("h");
            cs.AddCommand("B");
            cs.AddCommand("Q");
        }

        private static void Circle(PdfContentStream cs, string fillHex, string strokeHex)
        {
            cs.AddCommand(PdfColor.SetColorFromHex(fillHex).ToFillCommand());
            cs.AddCommand(PdfColor.SetColorFromHex(strokeHex).ToStrokeCommand());
            cs.AddCommand("6 w");
            cs.AddCommand("93 50 m");                        // circle cx50 cy50 r43 -> 4 cubics (k·r≈23.748)
            cs.AddCommand("93 73.748 73.748 93 50 93 c");
            cs.AddCommand("26.252 93 7 73.748 7 50 c");
            cs.AddCommand("7 26.252 26.252 7 50 7 c");
            cs.AddCommand("73.748 7 93 26.252 93 50 c");
            cs.AddCommand("h");
            cs.AddCommand("B");
        }

        private static void Arrow(PdfContentStream cs, double svgRotation, string fillHex, string strokeHex)
        {
            cs.AddCommand("q");
            if (svgRotation != 0d)
            {
                // The 0.8 scale + y-flip are baked into the path below; only rotation varies per arrow.
                // A baked y-flip reverses rotation sense, so the PDF angle is -svgRotation, about (50,50).
                double rad = -svgRotation * System.Math.PI / 180d;
                double cos = System.Math.Cos(rad), sin = System.Math.Sin(rad);
                double e = 50d * (1d - cos + sin);
                double f = 50d * (1d - sin - cos);
                cs.AddCommand($"{cos.ToPdfStringF4()} {sin.ToPdfStringF4()} {(-sin).ToPdfStringF4()} {cos.ToPdfStringF4()} {e.ToPdfStringF4()} {f.ToPdfStringF4()} cm");
            }
            cs.AddCommand(PdfColor.SetColorFromHex(fillHex).ToFillCommand());
            cs.AddCommand(PdfColor.SetColorFromHex(strokeHex).ToStrokeCommand());
            cs.AddCommand("2.56 w");                          // 3.2 · 0.8
            cs.AddCommand("34.423 89.473 m");                 // base points DOWN (scale 0.8 + flip baked)
            cs.AddCommand("34.422 41.330 l");
            cs.AddCommand("10.584 65.167 l");
            cs.AddCommand("10.526 43.283 l");
            cs.AddCommand("43.266 10.546 l");                 // tip
            cs.AddCommand("76.123 43.283 l");
            cs.AddCommand("76.123 65.167 l");
            cs.AddCommand("52.284 41.330 l");
            cs.AddCommand("52.284 89.474 l");
            cs.AddCommand("h");
            cs.AddCommand("B");
            cs.AddCommand("Q");
        }
    }
}
