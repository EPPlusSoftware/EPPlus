/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2025         EPPlus Software AB           Initial release EPPlus 8
 *************************************************************************************************/
using System.Collections.Generic;

namespace OfficeOpenXml.Drawing.Shape
{
    internal static class ShapeGuidesFactory
    {
        internal static Dictionary<string, ShapeGuidePoint> GetAdjustmentPoints(eShapeStyle style)
        {
            return style switch
            {
                eShapeStyle.BentConnector3 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.CurvedConnector3 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.RoundRect => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(25000 /*0-50000*/) }
                },
                eShapeStyle.Snip1Rect => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Snip2SameRect => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Snip2DiagRect => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.SnipRoundRect => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Round1Rect => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Round2SameRect => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Round2DiagRect => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Triangle => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(50000) /*0-1000000*/ }
                },
                eShapeStyle.Parallelogram => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Trapezoid => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Hexagon => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) },
                    { "vf", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Octagon => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Pie => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Chord => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Teardrop => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Frame => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.HalfFrame => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Corner => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.DiagStripe => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Plus => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Plaque => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Can => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Cube => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Bevel => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Donut => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.NoSmoking => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.BlockArc => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                },
                eShapeStyle.FoldedCorner => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.SmileyFace => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Sun => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Moon => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Arc => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.BracketPair => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.BracePair => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.LeftBracket => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.RightBracket => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.LeftBrace => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.RightBrace => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.RightArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.LeftArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.UpArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.DownArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.LeftRightArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.UpDownArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.QuadArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                },
                eShapeStyle.LeftRightUpArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                },
                eShapeStyle.BentArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                },
                eShapeStyle.UturnArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                    { "adj5", new ShapeGuidePoint(0) },
                },
                eShapeStyle.LeftUpArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                },
                eShapeStyle.BentUpArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                },
                eShapeStyle.CurvedRightArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                },
                eShapeStyle.CurvedLeftArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                },
                eShapeStyle.CurvedUpArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                },
                eShapeStyle.CurvedDownArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                },
                eShapeStyle.StripedRightArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.NotchedRightArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.HomePlate => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Chevron => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.RightArrowCallout => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                },
                eShapeStyle.DownArrowCallout => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                },
                eShapeStyle.LeftArrowCallout => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                },
                eShapeStyle.UpArrowCallout => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                },
                eShapeStyle.LeftRightArrowCallout => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                },
                eShapeStyle.QuadArrowCallout => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                },
                eShapeStyle.CircularArrow => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                    { "adj5", new ShapeGuidePoint(0) },
                },
                eShapeStyle.MathPlus => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.MathMinus => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.MathMultiply => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.MathDivide => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                },
                eShapeStyle.MathEqual => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.MathNotEqual => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Star4 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Star5 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) },
                    { "hf", new ShapeGuidePoint(0) },
                    { "vf", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Star6 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) },
                    { "hf", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Star7 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) },
                    { "hf", new ShapeGuidePoint(0) },
                    { "vf", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Star8 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Star10 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) },
                    { "hf", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Star12 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Star16 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Star24 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Star32 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Ribbon2 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Ribbon => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.EllipseRibbon2 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                },
                eShapeStyle.EllipseRibbon => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                },
                eShapeStyle.VerticalScroll => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.HorizontalScroll => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj", new ShapeGuidePoint(0) }
                },
                eShapeStyle.Wave => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.DoubleWave => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.WedgeRectCallout => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.WedgeRoundRectCallout => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                },
                eShapeStyle.WedgeEllipseCallout => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.CloudCallout => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                },
                eShapeStyle.BorderCallout1 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                },
                eShapeStyle.BorderCallout2 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                    { "adj5", new ShapeGuidePoint(0) },
                    { "adj6", new ShapeGuidePoint(0) },
                },
                eShapeStyle.BorderCallout3 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                    { "adj5", new ShapeGuidePoint(0) },
                    { "adj6", new ShapeGuidePoint(0) },
                    { "adj7", new ShapeGuidePoint(0) },
                    { "adj8", new ShapeGuidePoint(0) },
                },
                eShapeStyle.AccentCallout1 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                },
                eShapeStyle.AccentCallout2 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                    { "adj5", new ShapeGuidePoint(0) },
                    { "adj6", new ShapeGuidePoint(0) },
                },
                eShapeStyle.AccentCallout3 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                    { "adj5", new ShapeGuidePoint(0) },
                    { "adj6", new ShapeGuidePoint(0) },
                    { "adj7", new ShapeGuidePoint(0) },
                    { "adj8", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Callout1 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Callout2 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                    { "adj5", new ShapeGuidePoint(0) },
                    { "adj6", new ShapeGuidePoint(0) },
                },
                eShapeStyle.Callout3 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                    { "adj5", new ShapeGuidePoint(0) },
                    { "adj6", new ShapeGuidePoint(0) },
                    { "adj7", new ShapeGuidePoint(0) },
                    { "adj8", new ShapeGuidePoint(0) },
                },
                eShapeStyle.AccentBorderCallout1 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                },
                eShapeStyle.AccentBorderCallout2 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                    { "adj5", new ShapeGuidePoint(0) },
                    { "adj6", new ShapeGuidePoint(0) },
                },
                eShapeStyle.AccentBorderCallout3 => new Dictionary<string, ShapeGuidePoint>
                {
                    { "adj1", new ShapeGuidePoint(0) },
                    { "adj2", new ShapeGuidePoint(0) },
                    { "adj3", new ShapeGuidePoint(0) },
                    { "adj4", new ShapeGuidePoint(0) },
                    { "adj5", new ShapeGuidePoint(0) },
                    { "adj6", new ShapeGuidePoint(0) },
                    { "adj7", new ShapeGuidePoint(0) },
                    { "adj8", new ShapeGuidePoint(0) },
                },
                _ => null
            };
        }
    }
}