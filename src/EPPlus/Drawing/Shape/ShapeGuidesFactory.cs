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
using System.Linq;

namespace OfficeOpenXml.Drawing.Shape
{
    internal static class ShapeGuidesFactory
    {
        internal static Dictionary<eShapeStyle, Dictionary<string, ShapeGuidePoint>> DefaultAdjustments = new Dictionary<eShapeStyle, Dictionary<string, ShapeGuidePoint>>
        {
            {eShapeStyle.Parallelogram,new Dictionary<string, ShapeGuidePoint> { {"adj", 25000}}},
            {eShapeStyle.Trapezoid,new Dictionary<string, ShapeGuidePoint> { {"adj", 25000}}},
            {eShapeStyle.RoundRect,new Dictionary<string, ShapeGuidePoint> { {"adj", 16667}}},
            {eShapeStyle.Octagon,new Dictionary<string, ShapeGuidePoint> { {"adj", 29289}}},
            {eShapeStyle.Triangle,new Dictionary<string, ShapeGuidePoint> { {"adj", 50000}}},
            {eShapeStyle.Hexagon,new Dictionary<string, ShapeGuidePoint> { {"adj", 25000},{"vf", 115470}}},
            {eShapeStyle.Plus,new Dictionary<string, ShapeGuidePoint> { {"adj", 25000}}},
            {eShapeStyle.Pentagon,new Dictionary<string, ShapeGuidePoint> { {"hf", 105146},{"vf", 110557}}},
            {eShapeStyle.Can,new Dictionary<string, ShapeGuidePoint> { {"adj", 25000}}},
            {eShapeStyle.Cube,new Dictionary<string, ShapeGuidePoint> { {"adj", 25000}}},
            {eShapeStyle.Bevel,new Dictionary<string, ShapeGuidePoint> { {"adj", 12500}}},
            {eShapeStyle.FoldedCorner,new Dictionary<string, ShapeGuidePoint> { {"adj", 16667}}},
            {eShapeStyle.SmileyFace,new Dictionary<string, ShapeGuidePoint> { {"adj", 4653}}},
            {eShapeStyle.Donut,new Dictionary<string, ShapeGuidePoint> { {"adj", 25000}}},
            {eShapeStyle.NoSmoking,new Dictionary<string, ShapeGuidePoint> { {"adj", 18750}}},
            {eShapeStyle.BlockArc,new Dictionary<string, ShapeGuidePoint> { {"adj1", 18000000},{"adj2", 0},{"adj3", 25000}}},
            {eShapeStyle.Sun,new Dictionary<string, ShapeGuidePoint> { {"adj", 25000}}},
            {eShapeStyle.Moon,new Dictionary<string, ShapeGuidePoint> { {"adj", 50000}}},
            {eShapeStyle.Arc,new Dictionary<string, ShapeGuidePoint> { {"adj1", -9000000},{"adj2", 0}}},
            {eShapeStyle.BracketPair,new Dictionary<string, ShapeGuidePoint> { {"adj", 16667}}},
            {eShapeStyle.BracePair,new Dictionary<string, ShapeGuidePoint> { {"adj", 8333}}},
            {eShapeStyle.Plaque,new Dictionary<string, ShapeGuidePoint> { {"adj", 16667}}},
            {eShapeStyle.LeftBracket,new Dictionary<string, ShapeGuidePoint> { {"adj", 8333}}},
            {eShapeStyle.RightBracket,new Dictionary<string, ShapeGuidePoint> { {"adj", 8333}}},
            {eShapeStyle.LeftBrace,new Dictionary<string, ShapeGuidePoint> { {"adj1", 8333},{"adj2", 50000}}},
            {eShapeStyle.RightBrace,new Dictionary<string, ShapeGuidePoint> { {"adj1", 8333},{"adj2", 50000}}},
            {eShapeStyle.RightArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 50000},{"adj2", 50000}}},
            {eShapeStyle.LeftArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 50000},{"adj2", 50000}}},
            {eShapeStyle.UpArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 50000},{"adj2", 50000}}},
            {eShapeStyle.DownArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 50000},{"adj2", 50000}}},
            {eShapeStyle.LeftRightArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 50000},{"adj2", 50000}}},
            {eShapeStyle.UpDownArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 50000},{"adj2", 50000}}},
            {eShapeStyle.QuadArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 22500},{"adj2", 22500},{"adj3", 22500}}},
            {eShapeStyle.LeftRightUpArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 25000},{"adj3", 25000}}},
            {eShapeStyle.BentArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 25000},{"adj3", 25000},{"adj4", 43750}}},
            {eShapeStyle.UturnArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 25000},{"adj3", 25000},{"adj4", 43750},{"adj5", 75000}}},
            {eShapeStyle.LeftUpArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 25000},{"adj3", 25000}}},
            {eShapeStyle.BentUpArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 25000},{"adj3", 25000}}},
            {eShapeStyle.CurvedRightArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 50000},{"adj3", 25000}}},
            {eShapeStyle.CurvedLeftArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 50000},{"adj3", 25000}}},
            {eShapeStyle.CurvedUpArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 50000},{"adj3", 25000}}},
            {eShapeStyle.CurvedDownArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 50000},{"adj3", 25000}}},
            {eShapeStyle.StripedRightArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 50000},{"adj2", 50000}}},
            {eShapeStyle.NotchedRightArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 50000},{"adj2", 50000}}},
            {eShapeStyle.HomePlate,new Dictionary<string, ShapeGuidePoint> { {"adj", 50000}}},
            {eShapeStyle.Chevron,new Dictionary<string, ShapeGuidePoint> { {"adj", 50000}}},
            {eShapeStyle.RightArrowCallout,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 25000},{"adj3", 25000},{"adj4", 64977}}},
            {eShapeStyle.LeftArrowCallout,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 25000},{"adj3", 25000},{"adj4", 649767}}},
            {eShapeStyle.UpArrowCallout,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 25000},{"adj3", 25000},{"adj4", 649767}}},
            {eShapeStyle.DownArrowCallout,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 25000},{"adj3", 25000},{"adj4", 649767}}},
            {eShapeStyle.LeftRightArrowCallout,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 25000},{"adj3", 25000},{"adj4", 48123}}},
            {eShapeStyle.UpDownArrowCallout,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 25000},{"adj3", 25000},{"adj4", 48123}}},
            {eShapeStyle.QuadArrowCallout,new Dictionary<string, ShapeGuidePoint> { {"adj1", 18515},{"adj2", 18515},{"adj3", 18515},{"adj4", 48123}}},
            {eShapeStyle.CircularArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 12500},{"adj2", 1903865},{"adj3", -1903865},{"adj4", 18000000},{"adj5", 12500}}},
            {eShapeStyle.Star4,new Dictionary<string, ShapeGuidePoint> { {"adj", 12500}}},
            {eShapeStyle.Star5,new Dictionary<string, ShapeGuidePoint> { {"adj", 19098},{"hf", 105146},{"vf", 110557}}},
            {eShapeStyle.Star8,new Dictionary<string, ShapeGuidePoint> { {"adj", 37500}}},
            {eShapeStyle.Star16,new Dictionary<string, ShapeGuidePoint> { {"adj", 37500}}},
            {eShapeStyle.Star24,new Dictionary<string, ShapeGuidePoint> { {"adj", 37500}}},
            {eShapeStyle.Star32,new Dictionary<string, ShapeGuidePoint> { {"adj", 37500}}},
            {eShapeStyle.Ribbon2,new Dictionary<string, ShapeGuidePoint> { {"adj1", 16667},{"adj2", 50000}}},
            {eShapeStyle.Ribbon,new Dictionary<string, ShapeGuidePoint> { {"adj1", 16667},{"adj2", 50000}}},
            {eShapeStyle.EllipseRibbon2,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 50000},{"adj3", 12500}}},
            {eShapeStyle.EllipseRibbon,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 50000},{"adj3", 12500}}},
            {eShapeStyle.VerticalScroll,new Dictionary<string, ShapeGuidePoint> { {"adj", 12500}}},
            {eShapeStyle.HorizontalScroll,new Dictionary<string, ShapeGuidePoint> { {"adj", 12500}}},
            {eShapeStyle.Wave,new Dictionary<string, ShapeGuidePoint> { {"adj1", 12500},{"adj2", 0}}},
            {eShapeStyle.DoubleWave,new Dictionary<string, ShapeGuidePoint> { {"adj1", 6250},{"adj2", 0}}},
            {eShapeStyle.WedgeRectCallout,new Dictionary<string, ShapeGuidePoint> { {"adj1", -20833},{"adj2", 62500}}},
            {eShapeStyle.WedgeRoundRectCallout,new Dictionary<string, ShapeGuidePoint> { {"adj1", -20833},{"adj2", 62500},{"adj3", 16667},{"adj4", -20833},{"adj5", 62500},{"adj6", 16667}}},
            {eShapeStyle.WedgeEllipseCallout,new Dictionary<string, ShapeGuidePoint> { {"adj1", -20833},{"adj2", 62500}}},
            {eShapeStyle.CloudCallout,new Dictionary<string, ShapeGuidePoint> { {"adj1", -20833},{"adj2", 62500}}},
            {eShapeStyle.BorderCallout1,new Dictionary<string, ShapeGuidePoint> { {"adj1", 18750},{"adj2", -8333},{"adj3", 112500},{"adj4", -38333},{"adj5", 18750},{"adj6", -8333},{"adj7", 112500},{"adj8", -38333},{"adj9", 18750},{"adj10", -8333},{"adj11", 112500},{"adj12", -38333}}},
            {eShapeStyle.BorderCallout2,new Dictionary<string, ShapeGuidePoint> { {"adj1", 18750},{"adj2", -8333},{"adj3", 18750},{"adj4", -16667},{"adj5", 112500},{"adj6", -46667}}},
            {eShapeStyle.BorderCallout3,new Dictionary<string, ShapeGuidePoint> { {"adj1", 18750},{"adj2", -8333},{"adj3", 18750},{"adj4", -16667},{"adj5", 100000},{"adj6", -16667},{"adj7", 112963},{"adj8", -8333}}},
            {eShapeStyle.AccentCallout1,new Dictionary<string, ShapeGuidePoint> { {"adj1", 18750},{"adj2", -8333},{"adj3", 112500},{"adj4", -38333}}},
            {eShapeStyle.AccentCallout2,new Dictionary<string, ShapeGuidePoint> { {"adj1", 18750},{"adj2", -8333},{"adj3", 18750},{"adj4", -16667},{"adj5", 112500},{"adj6", -46667}}},
            {eShapeStyle.AccentCallout3,new Dictionary<string, ShapeGuidePoint> { {"adj1", 18750},{"adj2", -8333},{"adj3", 18750},{"adj4", -16667},{"adj5", 100000},{"adj6", -16667},{"adj7", 112963},{"adj8", -8333}}},
            {eShapeStyle.Callout1,new Dictionary<string, ShapeGuidePoint> { {"adj1", 18750},{"adj2", -8333},{"adj3", 112500},{"adj4", -38333},{"adj5", 18750},{"adj6", -8333},{"adj7", 112500},{"adj8", -38333},{"adj9", 18750},{"adj10", -8333},{"adj11", 112500},{"adj12", -38333}}},
            {eShapeStyle.Callout2,new Dictionary<string, ShapeGuidePoint> { {"adj1", 18750},{"adj2", -8333},{"adj3", 18750},{"adj4", -16667},{"adj5", 112500},{"adj6", -46667}}},
            {eShapeStyle.Callout3,new Dictionary<string, ShapeGuidePoint> { {"adj1", 18750},{"adj2", -8333},{"adj3", 18750},{"adj4", -16667},{"adj5", 100000},{"adj6", -16667},{"adj7", 112963},{"adj8", -8333}}},
            {eShapeStyle.AccentBorderCallout1,new Dictionary<string, ShapeGuidePoint> { {"adj1", 18750},{"adj2", -8333},{"adj3", 112500},{"adj4", -38333}}},
            {eShapeStyle.AccentBorderCallout2,new Dictionary<string, ShapeGuidePoint> { {"adj1", 18750},{"adj2", -8333},{"adj3", 18750},{"adj4", -16667},{"adj5", 112500},{"adj6", -46667}}},
            {eShapeStyle.AccentBorderCallout3,new Dictionary<string, ShapeGuidePoint> { {"adj1", 18750},{"adj2", -8333},{"adj3", 18750},{"adj4", -16667},{"adj5", 100000},{"adj6", -16667},{"adj7", 112963},{"adj8", -8333}}},
            {eShapeStyle.LeftRightRibbon,new Dictionary<string, ShapeGuidePoint> { {"adj1", 50000},{"adj2", 50000},{"adj3", 16667}}},
            {eShapeStyle.DiagStripe,new Dictionary<string, ShapeGuidePoint> { {"adj", 50000}}},
            {eShapeStyle.Pie,new Dictionary<string, ShapeGuidePoint> { {"adj1", 0},{"adj2", -9000000}}},
            {eShapeStyle.NonIsoscelesTrapezoid,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 25000}}},
            {eShapeStyle.Decagon,new Dictionary<string, ShapeGuidePoint> { {"vf", 105146}}},
            {eShapeStyle.Heptagon,new Dictionary<string, ShapeGuidePoint> { {"hf", 102572},{"vf", 105210}}},
            {eShapeStyle.Star6,new Dictionary<string, ShapeGuidePoint> { {"adj", 28868},{"hf", 115470}}},
            {eShapeStyle.Star7,new Dictionary<string, ShapeGuidePoint> { {"adj", 34601},{"hf", 102572},{"vf", 105210}}},
            {eShapeStyle.Star10,new Dictionary<string, ShapeGuidePoint> { {"adj", 42533},{"hf", 105146}}},
            {eShapeStyle.Star12,new Dictionary<string, ShapeGuidePoint> { {"adj", 37500}}},
            {eShapeStyle.Round1Rect,new Dictionary<string, ShapeGuidePoint> { {"adj", 16667}}},
            {eShapeStyle.Round2SameRect,new Dictionary<string, ShapeGuidePoint> { {"adj1", 16667},{"adj2", 0}}},
            {eShapeStyle.Round2DiagRect,new Dictionary<string, ShapeGuidePoint> { {"adj1", 16667},{"adj2", 0}}},
            {eShapeStyle.SnipRoundRect,new Dictionary<string, ShapeGuidePoint> { {"adj1", 16667},{"adj2", 16667}}},
            {eShapeStyle.Snip1Rect,new Dictionary<string, ShapeGuidePoint> { {"adj", 16667}}},
            {eShapeStyle.Snip2SameRect,new Dictionary<string, ShapeGuidePoint> { {"adj1", 16667},{"adj2", 0}}},
            {eShapeStyle.Snip2DiagRect,new Dictionary<string, ShapeGuidePoint> { {"adj1", 0},{"adj2", 16667}}},
            {eShapeStyle.Frame,new Dictionary<string, ShapeGuidePoint> { {"adj1", 12500}}},
            {eShapeStyle.HalfFrame,new Dictionary<string, ShapeGuidePoint> { {"adj1", 33333},{"adj2", 33333}}},
            {eShapeStyle.Teardrop,new Dictionary<string, ShapeGuidePoint> { {"adj", 100000}}},
            {eShapeStyle.Chord,new Dictionary<string, ShapeGuidePoint> { {"adj1", 4500000},{"adj2", -9000000}}},
            {eShapeStyle.Corner,new Dictionary<string, ShapeGuidePoint> { {"adj1", 50000},{"adj2", 50000}}},
            {eShapeStyle.MathPlus,new Dictionary<string, ShapeGuidePoint> { {"adj1", 23520}}},
            {eShapeStyle.MathMinus,new Dictionary<string, ShapeGuidePoint> { {"adj1", 23520}}},
            {eShapeStyle.MathMultiply,new Dictionary<string, ShapeGuidePoint> { {"adj1", 23520}}},
            {eShapeStyle.MathDivide,new Dictionary<string, ShapeGuidePoint> { {"adj1", 23520},{"adj2", 5880},{"adj3", 11760}}},
            {eShapeStyle.MathEqual,new Dictionary<string, ShapeGuidePoint> { {"adj1", 23520},{"adj2", 11760}}},
            {eShapeStyle.MathNotEqual,new Dictionary<string, ShapeGuidePoint> { {"adj1", 23520},{"adj2", 11000000},{"adj3", 11760}}},
            {eShapeStyle.Gear6,new Dictionary<string, ShapeGuidePoint> { {"adj1", 15000},{"adj2", 3526}}},
            {eShapeStyle.Gear9,new Dictionary<string, ShapeGuidePoint> { {"adj1", 10000},{"adj2", 1763}}},
            {eShapeStyle.LeftCircularArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 12500},{"adj2", -1903865},{"adj3", 1903865},{"adj4", 18000000},{"adj5", 12500}}},
            {eShapeStyle.LeftRightCircularArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 12500},{"adj2", 1903865},{"adj3", -1903865},{"adj4", -16096130},{"adj5", 12500}}},
            {eShapeStyle.SwooshArrow,new Dictionary<string, ShapeGuidePoint> { {"adj1", 25000},{"adj2", 16667}}},
        };

        internal static Dictionary<string, ShapeGuidePoint> GetAdjustmentPoints(eShapeStyle style)
        {
            if(DefaultAdjustments.TryGetValue(style, out Dictionary<string, ShapeGuidePoint> points))
            {
                return points;
            }
            return null;            
        }
        public static List<int> GetAdjustmentPointList(eShapeStyle style)
        {
            if (DefaultAdjustments.TryGetValue(style, out Dictionary<string, ShapeGuidePoint> points))
            {
                return points.Select(x => x.Value.Value).ToList(); ;
            }
            return null;
        }
        //return style switch
        //{
        //    eShapeStyle.BentConnector3 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.CurvedConnector3 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.RoundRect => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(25000 /*50000-50000*/) }
        //    },
        //    eShapeStyle.Snip1Rect => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Snip2SameRect => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Snip2DiagRect => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.SnipRoundRect => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Round1Rect => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Round2SameRect => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Round2DiagRect => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Triangle => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) /*50000-1000000*/ }
        //    },
        //    eShapeStyle.Parallelogram => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Trapezoid => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Hexagon => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) },
        //        { "vf", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Octagon => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Pie => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Chord => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Teardrop => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Frame => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.HalfFrame => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Corner => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.DiagStripe => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Plus => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Plaque => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Can => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Cube => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Bevel => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Donut => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.NoSmoking => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.BlockArc => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.FoldedCorner => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.SmileyFace => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Sun => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Moon => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Arc => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.BracketPair => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.BracePair => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.LeftBracket => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.RightBracket => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.LeftBrace => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.RightBrace => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.RightArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.LeftArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.UpArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.DownArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.LeftRightArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.UpDownArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.QuadArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.LeftRightUpArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.BentArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.UturnArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //        { "adj5", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.LeftUpArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.BentUpArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.CurvedRightArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.CurvedLeftArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.CurvedUpArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.CurvedDownArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.StripedRightArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.NotchedRightArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.HomePlate => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Chevron => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.RightArrowCallout => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.DownArrowCallout => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.LeftArrowCallout => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.UpArrowCallout => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.LeftRightArrowCallout => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.QuadArrowCallout => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.CircularArrow => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //        { "adj5", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.MathPlus => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.MathMinus => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.MathMultiply => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.MathDivide => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.MathEqual => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.MathNotEqual => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Star4 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Star5 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) },
        //        { "hf", new ShapeGuidePoint(50000) },
        //        { "vf", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Star6 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) },
        //        { "hf", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Star7 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) },
        //        { "hf", new ShapeGuidePoint(50000) },
        //        { "vf", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Star8 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Star10 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) },
        //        { "hf", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Star12 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Star16 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Star24 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Star32 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Ribbon2 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Ribbon => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.EllipseRibbon2 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.EllipseRibbon => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.VerticalScroll => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.HorizontalScroll => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj", new ShapeGuidePoint(50000) }
        //    },
        //    eShapeStyle.Wave => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.DoubleWave => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.WedgeRectCallout => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.WedgeRoundRectCallout => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.WedgeEllipseCallout => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.CloudCallout => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.BorderCallout1 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.BorderCallout2 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //        { "adj5", new ShapeGuidePoint(50000) },
        //        { "adj6", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.BorderCallout3 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //        { "adj5", new ShapeGuidePoint(50000) },
        //        { "adj6", new ShapeGuidePoint(50000) },
        //        { "adj7", new ShapeGuidePoint(50000) },
        //        { "adj8", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.AccentCallout1 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.AccentCallout2 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //        { "adj5", new ShapeGuidePoint(50000) },
        //        { "adj6", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.AccentCallout3 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //        { "adj5", new ShapeGuidePoint(50000) },
        //        { "adj6", new ShapeGuidePoint(50000) },
        //        { "adj7", new ShapeGuidePoint(50000) },
        //        { "adj8", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Callout1 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Callout2 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //        { "adj5", new ShapeGuidePoint(50000) },
        //        { "adj6", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.Callout3 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //        { "adj5", new ShapeGuidePoint(50000) },
        //        { "adj6", new ShapeGuidePoint(50000) },
        //        { "adj7", new ShapeGuidePoint(50000) },
        //        { "adj8", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.AccentBorderCallout1 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.AccentBorderCallout2 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //        { "adj5", new ShapeGuidePoint(50000) },
        //        { "adj6", new ShapeGuidePoint(50000) },
        //    },
        //    eShapeStyle.AccentBorderCallout3 => new Dictionary<string, ShapeGuidePoint>
        //    {
        //        { "adj1", new ShapeGuidePoint(50000) },
        //        { "adj2", new ShapeGuidePoint(50000) },
        //        { "adj3", new ShapeGuidePoint(50000) },
        //        { "adj4", new ShapeGuidePoint(50000) },
        //        { "adj5", new ShapeGuidePoint(50000) },
        //        { "adj6", new ShapeGuidePoint(50000) },
        //        { "adj7", new ShapeGuidePoint(50000) },
        //        { "adj8", new ShapeGuidePoint(50000) },
        //    },
        //    _ => null
        //};
    }
}
