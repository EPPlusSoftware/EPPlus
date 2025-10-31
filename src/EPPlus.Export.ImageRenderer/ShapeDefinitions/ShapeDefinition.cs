/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;

namespace EPPlusImageRenderer.ShapeDefinitions
{
    [DebuggerDisplay("{Style}")]
    internal class ShapeDefinition
    {
        internal Dictionary<string, double> _calculatedValues = new Dictionary<string, double>();
        internal Coordinate translateCoordinate = null;

        internal ShapeDefinition()
        {
                
        }
        /// <summary>
        /// Clone constructor
        /// </summary>
        /// <param name="original">The original to clone from</param>
        public ShapeDefinition(ShapeDefinition original)
        {
            Style=original.Style;
            if (original.ShapeAdjustValues != null)
            {
                ShapeAdjustValues = new List<ShapeGuide>();
                foreach (var av in original.ShapeAdjustValues)
                {
                    ShapeAdjustValues.Add(av.Clone());
                }
            }
            if (original.ShapeGuides != null)
            {
                ShapeGuides = new List<ShapeGuide>();
                foreach (var g in original.ShapeGuides)
                {
                    ShapeGuides.Add(g.Clone());
                }
            }

            if (original.ShapeAdjustHandles!=null)
            {
                ShapeAdjustHandles = new List<ShapeAdjustHandleBase>();
                foreach (var ah in original.ShapeAdjustHandles)
                {
                    ShapeAdjustHandles.Add(ah.Clone());
                }
            }

            TextBoxRect = original.TextBoxRect?.Clone();

            ShapePaths = new List<DrawingPath>();
            foreach (var p in original.ShapePaths)
            {
                ShapePaths.Add(p.Clone());
            }
            if (original.translateCoordinate != null)
            {
                translateCoordinate = original.translateCoordinate;
            }
        }

        internal string GetTransform(double rotation)
        {
            if(translateCoordinate == null && rotation==0)
            {
                return "";
            }
            var transform = "";
            if(translateCoordinate!=null)
            {
                transform = $"translate({translateCoordinate.X},{translateCoordinate.Y})";
            }
            if(rotation!=0)
            {
                if (string.IsNullOrEmpty(transform) == false)
                {
                    transform += " ";
                }
                transform = $"rotate({rotation.ToString(CultureInfo.InvariantCulture)})";
            }
            return $"transform=\"{transform}\"";
        }

        public eShapeStyle Style { get; set; }
        /// <summary>
        /// avLst
        /// </summary>
        public List<ShapeGuide> ShapeAdjustValues { get; set; }
        /// <summary>
        /// gdLst  
        /// </summary>
        public List<ShapeGuide> ShapeGuides { get; set; }
        /// <summary>
        /// ahLst 
        /// </summary>
        public List<ShapeAdjustHandleBase> ShapeAdjustHandles { get; set; }
        //cxnLst 
        public List<ShapeConnectionSite> ShapeConnectionSite { get; set; }
        //rect 
        /// <summary>
        /// The rectangle for the text inside the shape.
        /// </summary>
        public TextBoxRect TextBoxRect { get; set; }
        //pathLst
        /// <summary>
        /// Paths to draw the shape
        /// </summary>
        public List<DrawingPath> ShapePaths { get; set; }
        public void Calculate(ExcelShape shape)
        {
            InitCalculatedValues(shape);

            if (ShapeAdjustValues != null)
            {
                if (shape.HasCustomAdjustmentPoints())
                {
                    var names = shape.GetAdjustmentPointsNames();
                    var l = shape.GetAdjustmentPointsList();
                    for(int i=0;i<ShapeAdjustValues.Count;i++)
                    {
                        _calculatedValues.Add(names[i], Convert.ToDouble(l[i]));
                    }
                }
                else
                {
                    foreach (var ap in ShapeAdjustValues)
                    {
                        if (string.IsNullOrEmpty(ap.Formula) == false)
                        {
                            ap.CalculatedValue = CalculateFormula(ap.Formula);
                            _calculatedValues.Add(ap.Name, ap.CalculatedValue);
                        }
                    }
                }
            }
            if (ShapeGuides != null)
            {
                foreach (var g in ShapeGuides)
                {
                    g.CalculatedValue = CalculateFormula(g.Formula);

                    if(g.CalculatedValue == double.MaxValue || g.CalculatedValue == double.PositiveInfinity || g.CalculatedValue == double.NegativeInfinity)
                    {
                        throw new Exception($"Double overflowed during calculation of {g.Name}");
                    }

                    if (_calculatedValues.ContainsKey(g.Name))
                    {
                        _calculatedValues[g.Name] = g.CalculatedValue;
                    }
                    else
                    {
                         _calculatedValues.Add(g.Name, g.CalculatedValue);
                    }
                }
            }
            if(ShapeAdjustHandles != null)
            {
                foreach(var h in ShapeAdjustHandles)
                {
                    switch(h.AhType)
                    {
                        case ShapeAdjustHandleType.XY:
                            var xyH = (ShapeAdjustHandleXY)h;
                            xyH.MinimumVerticalAdjustment = GetValueOfNameOrCalculateValue(xyH.MinimumVerticalAdjustment);
                            xyH.MaximumVerticalAdjustment = GetValueOfNameOrCalculateValue(xyH.MaximumVerticalAdjustment);
                            xyH.MinimumHorizontalAdjustment = GetValueOfNameOrCalculateValue(xyH.MinimumHorizontalAdjustment);
                            xyH.MaximumHorizontalAdjustment = GetValueOfNameOrCalculateValue(xyH.MaximumHorizontalAdjustment);
                            break;
                        case ShapeAdjustHandleType.Polar:
                            var polarH = (ShapeAdjustHandlePolar)h;
                            polarH.MinimumAngleAdjustment = GetValueOfNameOrCalculateValue(polarH.MinimumAngleAdjustment);
                            polarH.MaximumAngleAdjustment = GetValueOfNameOrCalculateValue(polarH.MaximumAngleAdjustment);
                            polarH.MinimumRadialAdjustment = GetValueOfNameOrCalculateValue(polarH.MinimumRadialAdjustment);
                            polarH.MaximumRadialAdjustment = GetValueOfNameOrCalculateValue(polarH.MaximumRadialAdjustment);
                            break;
                    }
                }
            }

            if (ShapePaths.Count > 0)
            {
                foreach(var item in ShapePaths)
                {
                    var shapeWidth = (double)shape._width * (double)ExcelDrawing.EMU_PER_PIXEL;
                    var shapeHeight = (double)shape._height * (double)ExcelDrawing.EMU_PER_PIXEL;

                    var widthRatio = item.Width.HasValue ? (double)shapeWidth / (double)item.Width : 1D;
                    var heightRatio = item.Height.HasValue ? (double)shapeWidth / (double)item.Height : 1D;
                    item.Width = shapeWidth;
                    item.Height = shapeHeight;

                    foreach (var p in item.Paths)
                    {
                        switch(p.Type)
                        {
                            case PathDrawingType.Close:
                                break;
                            case PathDrawingType.ArcTo:
                                var arc = (ArcTo)p;
                                if (string.IsNullOrEmpty(arc.WidthRadiusName) == false)
                                {
                                    arc.WidthRadius = _calculatedValues[arc.WidthRadiusName];
                                }
                                else
                                {
                                    arc.WidthRadius = (double)((arc.WidthRadius ?? 0D) * widthRatio);
                                }
                                if (string.IsNullOrEmpty(arc.HeightRadiusName) == false)
                                {
                                    arc.HeightRadius = _calculatedValues[arc.HeightRadiusName];
                                }
                                else
                                {
                                    arc.HeightRadius = (double)((arc.HeightRadius ?? 0D) * heightRatio);
                                }
                                if (string.IsNullOrEmpty(arc.StartAngleName) == false) arc.StartAngle = _calculatedValues[arc.StartAngleName];
                                if (string.IsNullOrEmpty(arc.SwingAngleName) == false) arc.SwingAngle = _calculatedValues[arc.SwingAngleName];
                                break;
                            default:
                                var pb = (PathWithCoordinates)p;
                                for(var i=0; i<pb.Coordinates.Count;i++)
                                {
                                    var c = pb.Coordinates[i];
                                    if (string.IsNullOrEmpty(c.XName) == false)
                                    {
                                        c.X = _calculatedValues[c.XName];
                                    }
                                    else
                                    {
                                        c.X = (double)((c.X ?? 0D) * widthRatio);
                                    }

                                    if (string.IsNullOrEmpty(c.YName) == false)
                                    {
                                        c.Y = _calculatedValues[c.YName];
                                    }
                                    else
                                    {
                                        c.Y = (double)((c.Y ?? 0D) * heightRatio);
                                    }
                                }
                                break;
                        }
                    }
                }
            }
            if(TextBoxRect != null)
            {
                TextBoxRect.LeftValue = GetValue(TextBoxRect.LeftName) / (double)ExcelDrawing.EMU_PER_PIXEL;
                TextBoxRect.RightValue = GetValue(TextBoxRect.RightName) / (double)ExcelDrawing.EMU_PER_PIXEL;
                TextBoxRect.TopValue = GetValue(TextBoxRect.TopName) / (double)ExcelDrawing.EMU_PER_PIXEL;
                TextBoxRect.BottomValue = GetValue(TextBoxRect.BottomName) / (double)ExcelDrawing.EMU_PER_PIXEL;
            }
        }

        private object GetValueOfNameOrCalculateValue(object value)
        {
            if (value is string s && _calculatedValues.ContainsKey(s))
            {
                return _calculatedValues[s];
            }
            return value;
        }

        private bool ValidateShapeAdjustmentBounds(int i)
        {
            throw new NotImplementedException();
        }

        private void InitCalculatedValues(ExcelShape shape)
        {
            //var adjustedHeight = shape._height - 5.5d;
            //var adjustedWidth = shape._width - 3.625;

            var adjustedHeight = shape._height;
            //var adjustedWidth = shape._width - 10.5d;
            var adjustedWidth = shape._width;


            var w = (double)(adjustedWidth * (double)ExcelDrawing.EMU_PER_PIXEL);
            var h = (double)(adjustedHeight * (double)ExcelDrawing.EMU_PER_PIXEL);
            //Longest side
            var ls = Math.Max(h, w);
            //Shortest side
            var ss = Math.Min(h, w);

            _calculatedValues = new Dictionary<string, double>()
            {
                {"t", 0d },
                {"l", 0d },
                {"w", w },
                {"r", w },
                {"h", h },
                {"b", h },
                {"hc", w/2d },
                {"vc", h/2d },
                {"ls", ls },
                {"ss", ss },
                {"3cd4", 16200000.0d},
                {"3cd8", 8100000.0d},
                {"5cd8", 13500000.0d},
                {"7cd8", 18900000.0d},
                {"cd2", 10800000.0d},
                {"cd4", 5400000.0d},
                {"cd8", 2700000.0d},
                {"hd2", h/2d},
                {"hd3", h/3d},
                {"hd4", h/4d},
                {"hd5", h/5d},
                {"hd6", h/6d},
                {"hd8", h/8d},
                {"wd2", w/2d},
                {"wd3", w/3d},
                {"wd4", w/4d},
                {"wd5", w/5d},
                {"wd6", w/6d},
                {"wd8", w/8d},
                {"wd10", w/10d},
                {"wd12", w/12d},
                {"wd16", w/16d},
                {"wd32", w/32d},
                {"ssd2", ss/2d },
                {"ssd4", ss/4d },
                {"ssd6", ss/6d },
                {"ssd8", ss/8d },
                {"ssd16", ss/16d },
                {"ssd32", ss/32d },
            };
        }

        internal double CalculateFormula(string formula)
        {
            var tokens = formula.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var t = tokens[0];
            switch (t)
            {
                case "val":
                    var val = GetValue(tokens[1]);
                    return GetValue(tokens[1]);
                case "*/":
                    var divBy = GetValue(tokens[3]);
                    if (divBy == 0) return 0;
                    var val1 = GetValue(tokens[1]);
                    var val2 = GetValue(tokens[2]);

                    return (val1 * val2) / (double)divBy;
                case "+-":
                    return (GetValue(tokens[1]) + GetValue(tokens[2])) - GetValue(tokens[3]);
                case "+/":
                    divBy = GetValue(tokens[3]);
                    if (divBy == 0) return 0;
                    return (GetValue(tokens[1]) + GetValue(tokens[2])) / (double)divBy;
                case "?:":
                    return GetValue(tokens[1]) > 0 ? GetValue(tokens[2]) : GetValue(tokens[3]);
                case "sqrt":
                    return Math.Sqrt(Math.Abs(GetValue(tokens[1])));
                case "abs":
                    return Math.Abs(GetValue(tokens[1]));
                case "min":
                    return Math.Min(GetValue(tokens[1]), GetValue(tokens[2]));
                case "max":
                    return Math.Max(GetValue(tokens[1]), GetValue(tokens[2]));
                case "mod":
                    return Math.Sqrt(Math.Pow((double)GetValue(tokens[1]), 2d) + Math.Pow((double)GetValue(tokens[2]), 2) + Math.Pow((double)GetValue(tokens[3]), 2));
                case "pin":
                    //if (y < x), then x = value of this guide else if (y > z), then z
                    double x = GetValue(tokens[1]);
                    double y = GetValue(tokens[2]);
                    double z = GetValue(tokens[3]);

                    return (y < x ? x : y > z ? z : y);
                case "sin":
                    var angleSin = GetValue(tokens[2]) / 60000D;
                    if(angleSin == 0d || angleSin == 180d)
                    {
                        return 0;
                    }
                    var radAngleSin = Math.Sin(MathHelper.Radians(angleSin));
                    return (GetValue(tokens[1]) * radAngleSin);
                case "cos":
                    var angleCos = GetValue(tokens[2]) / 60000D;

                    if (angleCos == 90d || angleCos == 270d)
                    {
                        return 0;
                    }
                    var radAngleCos = Math.Cos(MathHelper.Radians(angleCos));
                    return (GetValue(tokens[1]) * radAngleCos);
                case "tan":
                    var angleTan = GetValue(tokens[2]) / 60000D;

                    if (angleTan == 0d || angleTan == 90d)
                    {
                        //Tan technically undefined
                        return 0;
                    }

                    return (GetValue(tokens[1]) * Math.Tan(MathHelper.Radians(angleTan)));
                case "at2":
                    x = GetValue(tokens[1]);
                    y = GetValue(tokens[2]);
                    double angleRad = Math.Atan2(y, x);     // radians
                    double angleDeg = angleRad * (180d / Math.PI); // degrees

                    while (angleDeg < -360)
                    {
                        angleDeg += 360;                        // normalize to [0, 360)
                    }
                    while (angleDeg > 360)
                    {
                        angleDeg -= 360;
                    }

                    return (angleDeg * 60000D);
                case "cat2":
                    x = GetValue(tokens[1]);
                    y = GetValue(tokens[2]);
                    z = GetValue(tokens[3]);

                    double angleRadCatOrig = Math.Atan2(z, y);     // radians
                    double angleDegCat2Orig = angleRadCatOrig * (180.0 / Math.PI); // degrees

                    if (angleDegCat2Orig == 90d || angleDegCat2Orig == 270d)
                    {
                        return 0;
                    }

                    var dist = (x * Math.Cos(angleRadCatOrig));
                    return dist;
                case "sat2":
                    x = GetValue(tokens[1]);
                    y = GetValue(tokens[2]);
                    z = GetValue(tokens[3]);

                    double angleRadSatOrig = Math.Atan2(z, y);     // radians
                    double angleDegSat2Orig = angleRadSatOrig * (180.0 / Math.PI); // degrees

                    if (angleDegSat2Orig == 0d || angleDegSat2Orig == 180d)
                    {
                        return 0;
                    }

                    var sat2Rad = (x * Math.Sin(angleRadSatOrig));

                    return sat2Rad;
                default:
                    if (_calculatedValues.TryGetValue(t, out var v))
                    {
                        return v;
                    }
                    throw new InvalidOperationException($"Unknown function or variable {{{t}}}");
            }

        }
        private double GetValue(string v)
        {
            if(double.TryParse(v, out var l))
            {
                return l;
            }
            else
            {
                if(_calculatedValues.TryGetValue(v, out var cv))
                {
                    return cv;
                }
                throw new InvalidOperationException($"Unknown variable {{{v}}}");
            }
        }

        internal ShapeDefinition Clone()
        {
            return new ShapeDefinition(this);
        }
    }
}