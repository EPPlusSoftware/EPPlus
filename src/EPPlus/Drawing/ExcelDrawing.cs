/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System.Collections.Generic;
using OfficeOpenXml.Core.Worksheet;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Drawing.OleObject;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils.EnumUtils;
using OfficeOpenXml.Utils.FileUtils;
using OfficeOpenXml.Utils.XML;
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.Utils.TypeConversion;

namespace OfficeOpenXml.Drawing
{
    public enum PathDrawingType
    {
        MoveTo,
        LineTo,
        ArcTo,
        CubicBezTo,
        QuadBezerTo,
        Close
    }
    /// <summary>
    /// How a shape path is filled.
    /// </summary>
    public enum PathFillMode
    {
        /// <summary>
        /// The corresponding path should have a normally shaded color applied to it’s fill
        /// </summary>
        Norm,
        /// <summary>
        /// The corresponding path should have a darker shaded color applied to it’s fill.
        /// </summary>
        Darken,
        /// <summary>
        /// The corresponding path should have a slightly darker shaded color applied to it’s fill.
        /// </summary>
        DarkenLess,
        /// <summary>
        /// The corresponding path should have a lightly shaded color applied to it’s fill.
        /// </summary>
        Lighten,
        /// <summary>
        /// The corresponding path should have a slightly lighter shaded color applied to it’s fill.
        /// </summary>
        LightenLess,
        /// <summary>
        /// The corresponding path should have no fill.
        /// </summary>
        None
    }
    internal class DrawCoordinate
    {
        public DrawCoordinate(DrawCoordinate c) 
        {
            X = c.X;
            Y = c.Y;
            XName = c.XName;
            YName = c.YName;
        }

        public DrawCoordinate(object x, object y)
        {
            if(x is long xl)
            {
                X = xl;
            }
            else
            {
                XName = x.ToString();
                X = null;
            }
            if (y is long yl)
            {
                Y = yl;
            }
            else
            {
                YName = y.ToString();
                Y = null;
            }

        }
        public double? X { get; set; }
        public double? Y { get; set; }
        public string XName { get; set; }
        public string YName { get; set; }
    }
    public abstract class PathsBase
    {
        public abstract PathDrawingType Type { get;  }

        internal abstract PathsBase Clone();
        public abstract double EndX { get; }
        public abstract double EndY { get; }
    }
    internal abstract class PathWithCoordinates : PathsBase
    {
        protected PathWithCoordinates(XmlElement e)
        {
            foreach (var cn in e.ChildNodes)
            {
                if (cn is XmlElement ce && ce.LocalName=="pt")
                {
                    Coordinates.Add(new DrawCoordinate(GetNameOrNumber(ce.GetAttribute("x")), GetNameOrNumber(ce.GetAttribute("y"))));
                }
            }
        }

        private object GetNameOrNumber(string s)
        {
            if(long.TryParse(s, NumberStyles.Number , CultureInfo.InvariantCulture, out var l))
            {
                return l;
            }
            return s;
        }

        protected PathWithCoordinates(XmlReader xr)
        {
            var name = xr.LocalName;
            while(xr.Read())
            {
                if (xr.LocalName == "pt" && xr.NodeType == XmlNodeType.Element)
                {
                    Coordinates.Add(new DrawCoordinate(GetNameOrNumber(xr.GetAttribute("x")), GetNameOrNumber(xr.GetAttribute("y"))));
                }
                else if(xr.IsEndElementWithName(name))
                {
                    break;
                }
            }
        }

        protected PathWithCoordinates(PathWithCoordinates clone) 
        {
            foreach(var c in clone.Coordinates)
            {
                Coordinates.Add(new DrawCoordinate(c));
            }
        }
        public List<DrawCoordinate> Coordinates { get; set; } = new List<DrawCoordinate>();
        public override double EndX => Coordinates.Count > 0D ? Coordinates[Coordinates.Count-1].X.Value : 0D;
        public override double EndY => Coordinates.Count > 0D ? Coordinates[Coordinates.Count - 1].Y.Value : 0D;
    }
    internal class MoveTo : PathWithCoordinates
    {
        public MoveTo(MoveTo clone) : base(clone)
        {

        }
        public MoveTo(XmlElement e) : base(e)
        {
        }
        public MoveTo(XmlReader xr) : base(xr)
        {
        }
        public override PathDrawingType Type => PathDrawingType.MoveTo;
        public DrawCoordinate Coordinate { get; set; }

        internal override PathsBase Clone()
        {
            return new MoveTo(this);
        }
    }
    internal class LineTo : PathWithCoordinates
    {
        public LineTo(LineTo clone) : base(clone)
        {

        }
        public LineTo(XmlReader xr) : base(xr)
        {

        }

        public LineTo(XmlElement e):base(e)
        {

        }
        public override PathDrawingType Type => PathDrawingType.LineTo;
        public DrawCoordinate Coordinate { get; set; }
        internal override PathsBase Clone()
        {
            return new LineTo(this);
        }
    }
    internal class ClosePath : PathsBase
    {
        public ClosePath()
        {
            
        }
        public override PathDrawingType Type => PathDrawingType.Close;
        internal override PathsBase Clone()
        {
            return new ClosePath();
        }
        public override double EndX => double.MinValue;
        public override double EndY => double.MinValue;
    }
    internal class QuadBezerTo : PathWithCoordinates
    {
        public QuadBezerTo(QuadBezerTo clone) : base(clone)
        {

        }
        public QuadBezerTo(XmlReader xr) : base(xr)
        {

        }
        public QuadBezerTo(XmlElement e) : base(e)
        {
            
        }
        public override PathDrawingType Type => PathDrawingType.QuadBezerTo;
        internal override PathsBase Clone()
        {
            return new QuadBezerTo(this);
        }

    }

    internal class CubicBezerTo : PathWithCoordinates
    {
        public CubicBezerTo(CubicBezerTo clone) : base(clone)
        {
            
        }
        public CubicBezerTo(XmlReader xr) : base(xr)
        {

        }
        public CubicBezerTo(XmlElement e) : base(e)
        {

        }

        public override PathDrawingType Type => PathDrawingType.CubicBezTo;
        internal override PathsBase Clone()
        {
            return new CubicBezerTo(this);
        }
    }

    internal class  ArcTo : PathsBase
    {
        public ArcTo(XmlReader xr)
        {
            if (long.TryParse(xr.GetAttribute("hR"), out var hrv))
            {
                HeightRadius = hrv;
            }
            else
            {
                HeightRadiusName = xr.GetAttribute("hR");
            }

            if (long.TryParse(xr.GetAttribute("wR"), out var wrv))
            {
                WidthRadius = wrv;
            }
            else
            {
                WidthRadiusName = xr.GetAttribute("wR");
            }

            if (long.TryParse(xr.GetAttribute("swAng"), out var swAng))
            {
                SwingAngle = swAng;
            }
            else
            {
                SwingAngleName = xr.GetAttribute("swAng");
            }

            if (long.TryParse(xr.GetAttribute("stAng"), out var stAng))
            {
                StartAngle = stAng;
            }
            else
            {
                StartAngleName = xr.GetAttribute("stAng");
            }
        }
        public ArcTo(XmlElement e)
        {
            if(long.TryParse(e.GetAttribute("hR"), out var hrv))
            {
                HeightRadius = hrv;
            }
            else
            {
                HeightRadiusName = e.GetAttribute("hR");
            }

            if (long.TryParse(e.GetAttribute("wR"), out var wrv))
            {
                WidthRadius = wrv;
            }
            else
            {
                WidthRadiusName = e.GetAttribute("wR");
            }

            if (long.TryParse(e.GetAttribute("swAng"), out var swAng))
            {
                SwingAngle = swAng;
            }
            else
            {
                SwingAngleName = e.GetAttribute("swAng");
            }

            if (long.TryParse(e.GetAttribute("stAng"), out var stAng))
            {
                StartAngle = stAng;
            }
            else
            {
                StartAngleName = e.GetAttribute("stAng");
            }
        }
        public override PathDrawingType Type => PathDrawingType.ArcTo;
        public double? HeightRadius { get; set; }
        public double? StartAngle { get; set; }
        public double? SwingAngle { get; set; }
        public double? WidthRadius { get; set; }
        public string HeightRadiusName { get; set; }
        public string StartAngleName { get; set; }
        public string SwingAngleName { get; set; }
        public string WidthRadiusName { get; set; }
        private ArcTo()
        {
            
        }
        internal override PathsBase Clone()
        {
            return new ArcTo()
            {
                HeightRadius = HeightRadius,
                StartAngle = StartAngle,
                SwingAngle = SwingAngle,
                WidthRadius = WidthRadius,
                HeightRadiusName = HeightRadiusName,
                StartAngleName = StartAngleName,
                SwingAngleName = SwingAngleName,
                WidthRadiusName = WidthRadiusName
            };
        }
        double _endX, _endY;
        internal void SetEndCoordinates(double x, double y)
        {
            _endX = x;
            _endY = y;
        }
        public override double EndX => _endX;
        public override double EndY => _endY;
    }
    internal class DrawingPath
    {
        public DrawingPath(DrawingPath clone)
        {
            Width = clone.Width;
            Height = clone.Height;
            Fill = clone.Fill;
            Stroke = clone.Stroke;
            ExtrusionOk = clone.ExtrusionOk;
            foreach(var p in clone.Paths)
            {
                Paths.Add(p.Clone());
            }
        }
        public DrawingPath(XmlReader xr)
        {
            Width = ConvertUtil.GetValueLongNull(xr.GetAttribute("w"));
            Height = ConvertUtil.GetValueLongNull(xr.GetAttribute("h"));
            Fill = GetFill(xr.GetAttribute("fill"));
            Stroke = ConvertUtil.ToBooleanString(xr.GetAttribute("stroke"), true);
            ExtrusionOk = ConvertUtil.ToBooleanString(xr.GetAttribute("extrusionOk"), false);
            while (xr.Read())
            {
                if (xr.NodeType == XmlNodeType.Element)
                {
                    switch (xr.LocalName)
                    {
                        case "moveTo":
                            Paths.Add(new MoveTo(xr));
                            break;
                        case "lnTo":
                            Paths.Add(new LineTo(xr));
                            break;
                        case "cubicBezTo":
                            Paths.Add(new CubicBezerTo(xr));
                            break;
                        case "quadBezTo":
                            Paths.Add(new QuadBezerTo(xr));
                            break;
                        case "arcTo":
                            Paths.Add(new ArcTo(xr));
                            break;
                        case "close":
                            Paths.Add(new ClosePath());
                            break;
                    }
                }
                else if(xr.LocalName =="path" && xr.NodeType == XmlNodeType.EndElement)
                {
                    break;
                }
            }
        }

        public DrawingPath(XmlElement topNode, XmlNamespaceManager nsm)
        {
            Width = int.Parse(topNode.GetAttribute("w"));
            Height = int.Parse(topNode.GetAttribute("h"));
            Fill = GetFill(topNode.GetAttribute("fill"));
            Stroke = ConvertUtil.ToBooleanString(topNode.GetAttribute("stroke"), true);
            ExtrusionOk = ConvertUtil.ToBooleanString(topNode.GetAttribute("extrusionOk"), true);
            foreach (var child in topNode.ChildNodes)
            {
                if (child is XmlElement e)
                {
                    switch (e.LocalName)
                    {
                        case "moveTo":
                            Paths.Add(new MoveTo(e));
                            break;
                        case "lnTo":
                            Paths.Add(new LineTo(e));
                            break;
                        case "cubicBezTo":
                            Paths.Add(new CubicBezerTo(e));
                            break;
                        case "quadBezTo":
                            Paths.Add(new CubicBezerTo(e));
                            break;
                        case "arcTo":
                            Paths.Add(new ArcTo(e));
                            break;
                        case "close":
                            Paths.Add(new ClosePath());
                            break;
                    }
                }
            }
        }

        private PathFillMode GetFill(string s)
        {
            if (string.IsNullOrEmpty(s) == false)
            {
                return (PathFillMode)Enum.Parse(typeof(PathFillMode), s, true);
            }
            return PathFillMode.Norm;
        }

        internal DrawingPath Clone() => new DrawingPath(this);

        public bool Stroke { get; set; }
        public bool ExtrusionOk { get; set; }        
        public PathFillMode Fill { get; set; }
        public double? Width { get; set; }
        public double? Height { get; set; }
        public List<PathsBase> Paths { get; set; } = new List<PathsBase>();
    }
    /// <summary>
    /// Base class for drawings. 
    /// Drawings are Charts, Shapes and Pictures.
    /// </summary>
    public class ExcelDrawing : XmlHelper, IDisposable
    {
        internal ExcelDrawings _drawings;
        internal ExcelGroupShape _parent;
        internal string _topPath, _nvPrPath, _hyperLinkPath;
        internal string _topPathUngrouped, _nvPrPathUngrouped;
        internal int _id;
        internal const float STANDARD_DPI = 96;
        /// <summary>
        /// The ratio between EMU and Pixels
        /// </summary>
        public const int EMU_PER_PIXEL = 9525;
        /// <summary>
        /// The ratio between EMU and Points
        /// </summary>
        public const int EMU_PER_POINT = 12700;
        /// <summary>
        /// The ratio between EMU and centimeters
        /// </summary>
        public const int EMU_PER_CM = 360000;
        /// <summary>
        /// The ratio between EMU and millimeters
        /// </summary>
        public const int EMU_PER_MM = 3600000;
        /// <summary>
        /// The ratio between EMU and US Inches
        /// </summary>
        public const int EMU_PER_US_INCH = 914400;
        /// <summary>
        /// The ratio between EMU and pica
        /// </summary>
        public const int EMU_PER_PICA = EMU_PER_US_INCH / 6;

        internal double _width = double.MinValue, _height = double.MinValue, _top = double.MinValue, _left = double.MinValue;
        internal static readonly string[] _schemaNodeOrderSpPr = new string[] { "xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "pattFill", "grpFill", "blipFill", "ln", "effectLst", "effectDag", "scene3d", "sp3d" };

        internal bool _doNotAdjust = false;

        internal static string[] NamespacePrefixes = { "xdr", "cdr" };
        internal readonly int _prefixIndex = 0;
        internal readonly DrawingsCollectionType _collectionType;
        internal ExcelDrawingCoordinate _frmXPosition;
        internal ExcelDrawingSize _frmXSize;
        internal ExcelDrawing(ExcelDrawings drawings, XmlNode node, string topPath, string nvPrPath, ExcelGroupShape parent = null, DrawingsCollectionType collectionType = DrawingsCollectionType.Worksheet) :
            base(drawings.NameSpaceManager, node)
        {
            _drawings = drawings;
            _parent = parent;
            _prefixIndex = (int)collectionType;
            _collectionType = collectionType;
            TopNode = node;
            AddSchemaNodeOrder(new string[] { "from", "pos", "to", "ext", "pic", "graphicFrame", "sp", "cxnSp ", "grpSp", "nvSpPr", "nvCxnSpPr", "nvGraphicFramePr", "spPr", "style", "AlternateContent", "clientData" }, _schemaNodeOrderSpPr);
            _topPathUngrouped = topPath;
            _nvPrPathUngrouped = nvPrPath;

            if (_parent == null)
            {
                AdjustXPathsForGrouping(false);
                CellAnchor = GetAnchorFromName(node.LocalName);
                SetPositionProperties(drawings, node);
                GetPositionSize();          //Get the drawing position and size, so we can adjust it upon save, if the normal font is changed 

                string relID = GetXmlNodeString(_hyperLinkPath + "/@r:id");
                if (!string.IsNullOrEmpty(relID))
                {
                    HypRel = drawings.Part.GetRelationship(relID);

                    if (HypRel.TargetUri == null)
                    {
                        if (!string.IsNullOrEmpty(HypRel.Target))
                        {
                            _hyperLink = new ExcelHyperLink(HypRel.Target.Substring(1), "");
                        }
                    }
                    else
                    {
                        if (HypRel.TargetUri.IsAbsoluteUri)
                        {
                            _hyperLink = new ExcelHyperLink(HypRel.TargetUri.AbsoluteUri);
                        }
                        else
                        {
                            _hyperLink = new ExcelHyperLink(HypRel.TargetUri.OriginalString, UriKind.Relative);
                        }
                    }
                    if (Hyperlink is ExcelHyperLink ehl)
                    {
                        ehl.ToolTip = GetXmlNodeString(_hyperLinkPath + "/@tooltip");
                    }
                }
            }
            else
            {
                AdjustXPathsForGrouping(true);
                SetPositionProperties(drawings, node);
                GetPositionSize();                                  //Get the drawing position and size, so we can adjust it upon save, if the normal font is changed 
            }
            var custGeomNode = GetNode("xdr:sp/xdr:spPr/a:custGeom");
            if(custGeomNode!=null)
            {
                CustomGeom = new ExcelDrawingCustomGeometry(this, NameSpaceManager, custGeomNode);
            }            
            if (DrawingType == eDrawingType.Control || DrawingType == eDrawingType.OleObject || drawings._nextDrawingId >= 1025)
            {
                _id = drawings.Worksheet._nextControlId++;
            }            
        }

        internal virtual void AdjustXPathsForGrouping(bool group)
        {
            if (group)
            {
                _topPath = _topPathUngrouped.IndexOf('/') > 0 ? _topPathUngrouped.Substring(_topPathUngrouped.IndexOf('/') + 1) : "";
                if (_topPath == "")
                {
                    _nvPrPath = _nvPrPathUngrouped;
                }
                else
                {
                    _nvPrPath = _topPath + "/" + _nvPrPathUngrouped;
                }
            }
            else
            {
                _topPath = _topPathUngrouped;
                _nvPrPath = _topPath + "/" + _nvPrPathUngrouped;
            }
            _hyperLinkPath = $"{_nvPrPath}/a:hlinkClick";
        }

        internal void SetGroupChild(XmlNode offNode, XmlNode extNode)
        {
            CellAnchor = eEditAs.Absolute;

            From = null;
            To = null;
            Position = new ExcelDrawingCoordinate(NameSpaceManager, offNode, GetPositionSize);
            Size = new ExcelDrawingSize(NameSpaceManager, extNode, GetPositionSize);
        }

        internal void RemoveFromToNodes()
        {
            CellAnchor = eEditAs.Absolute;
            From = null;
            To = null;
        }

        private void SetPositionProperties(ExcelDrawings drawings, XmlNode node)
        {
            if (_parent == null) //Top level drawing
            {
                SetPositionPropertiesTopDrawing(drawings, node);
            }
            else //Child to Group shape
            {
                From = null;
                To = null;
                XmlNode posNode = GetXFrameNode(node, "a:off");
                if (posNode != null)
                {
                    Position = new ExcelDrawingCoordinate(drawings.NameSpaceManager, posNode, GetPositionSize);
                }

                posNode = GetXFrameNode(node, "a:ext");
                if (posNode != null)
                {
                    Size = new ExcelDrawingSize(drawings.NameSpaceManager, posNode, GetPositionSize);
                }
            }
            }

        private void SetPositionPropertiesTopDrawing(ExcelDrawings drawings, XmlNode node)
        {
            XmlNode posNode = node.SelectSingleNode(NamespacePrefixes[_prefixIndex] + ":from", drawings.NameSpaceManager);
            if (posNode != null)
            {
                From = new ExcelPosition(drawings.NameSpaceManager, posNode, GetPositionSize, _prefixIndex);
                Position = null;
            }
            else
            {
                posNode = node.SelectSingleNode("xdr:pos", drawings.NameSpaceManager);
                if (posNode != null)
                {
                    Position = new ExcelDrawingCoordinate(drawings.NameSpaceManager, posNode, GetPositionSize);
                }
            }
            posNode = node.SelectSingleNode(NamespacePrefixes[_prefixIndex] + ":to", drawings.NameSpaceManager);
            if (posNode != null)
            {
                To = new ExcelPosition(drawings.NameSpaceManager, posNode, GetPositionSize, _prefixIndex);
                Size = null;
            }
            else
            {
                To = null;
                posNode = node.SelectSingleNode(NamespacePrefixes[_prefixIndex] + ":ext", drawings.NameSpaceManager);
                if (posNode != null)
                {
                    Size = new ExcelDrawingSize(drawings.NameSpaceManager, posNode, GetPositionSize, _prefixIndex);
                }
            }
        }
        private XmlNode GetXFrameNode(XmlNode node, string child)
        {
            if (node.LocalName == "AlternateContent")
            {
                node = node.GetChildAtPosition(0).GetChildAtPosition(0);
            }
            if (node.LocalName == "grpSp")
            {
                return node.SelectSingleNode($"{NamespacePrefixes[_prefixIndex]}:grpSpPr/a:xfrm/{child}", NameSpaceManager);
            }
            else if (node.LocalName == "graphicFrame")
            {
                return node.SelectSingleNode($"{NamespacePrefixes[_prefixIndex]}:xfrm/{child}", NameSpaceManager);
            }
            else
            {
                return node.SelectSingleNode($"{NamespacePrefixes[_prefixIndex]}:spPr/a:xfrm/{child}", NameSpaceManager);
            }
        }

        internal bool IsWithinColumnRange(int colFrom, int colTo)
        {
            if (CellAnchor == eEditAs.OneCell)
            {

                GetToColumnFromPixels(_width, out int col, out _);
                return ((From.Column > colFrom - 1 || (From.Column == colFrom - 1 && From.ColumnOff == 0)) && (col <= colTo));
            }
            else if (CellAnchor == eEditAs.TwoCell)
            {
                return ((From.Column > colFrom - 1 || (From.Column == colFrom - 1 && From.ColumnOff == 0)) && (To.Column <= colTo));
            }
            else
            {
                return false;
            }
        }
        internal bool IsWithinRowRange(int rowFrom, int rowTo)
        {
            if (CellAnchor == eEditAs.OneCell)
            {
                GetToRowFromPixels(_height, out int row, out int pixOff);
                return ((From.Row > rowFrom - 1 || (From.Row == rowFrom - 1 && From.RowOff == 0)) && (row <= rowTo));
            }
            else if (CellAnchor == eEditAs.TwoCell)
            {
                return ((From.Row > rowFrom - 1 || (From.Row == rowFrom - 1 && From.RowOff == 0)) && (To.Row <= rowTo));
            }
            else
            {
                return false;
            }
        }

        internal static eEditAs GetAnchorFromName(string topElementName)
        {
            switch (topElementName)
            {
                case "oneCellAnchor":
                    return eEditAs.OneCell;
                case "absSizeAnchor":       //For drawings inside a chart 
                case "absoluteAnchor":
                    return eEditAs.Absolute;
                case "relSizeAnchor":       //For drawings inside a chart 
                case "twoCellAnchor":       
                default:
                    return eEditAs.TwoCell;
            }
        }
        /// <summary>
        /// The type of drawing
        /// </summary>
        public virtual eDrawingType DrawingType
        {
            get
            {
                return eDrawingType.Drawing;
            }
        }
        /// <summary>
        /// The name of the drawing object
        /// </summary>
        public virtual string Name
        {
            get
            {
                try
                {
                    if (_nvPrPath == "") return "";
                    return GetXmlNodeString(_nvPrPath + "/@name");
                }
                catch
                {
                    return "";
                }
            }
            set
            {
                try
                {
                    if (_nvPrPath == "") throw new NotImplementedException();
                    SetXmlNodeString(_nvPrPath + "/@name", value);
                    if (this is ExcelSlicer<ExcelTableSlicerCache> ts)
                    {
                        SetXmlNodeString(_nvPrPath + "/../../a:graphic/a:graphicData/sle:slicer/@name", value);
                        ts.SlicerName = value;
                    }
                    else if (this is ExcelSlicer<ExcelPivotTableSlicerCache> pts)
                    {
                        SetXmlNodeString(_nvPrPath + "/../../a:graphic/a:graphicData/sle:slicer/@name", value);
                        pts.SlicerName = value;
                    }
                }
                catch
                {
                    throw new NotImplementedException();
                }
            }
        }


        /// <summary>
        /// A description of the drawing object
        /// </summary>
        public string Description
        {
            get
            {
                try
                {
                    if (_nvPrPath == "") return "";
                    return GetXmlNodeString(_nvPrPath + "/@descr");
                }
                catch
                {
                    return "";
                }
            }
            set
            {
                try
                {
                    if (_nvPrPath == "") throw new NotImplementedException();
                    SetXmlNodeString(_nvPrPath + "/@descr", value);
                }
                catch
                {
                    throw new NotImplementedException();
                }
            }
        }
        /// <summary>
        /// How Excel resize drawings when the column width is changed within Excel.
        /// </summary>
        public eEditAs EditAs
        {
            get
            {
                try
                {
                    if (_parent != null && DrawingType == eDrawingType.Control)
                    {
                        return ((ExcelControl)this).GetCellAnchorFromWorksheetXml();
                    }
                    if (_parent != null && DrawingType == eDrawingType.OleObject)
                    {
                        return ((ExcelOleObject)this).GetCellAnchorFromWorksheetXml();
                    }
                    if (CellAnchor == eEditAs.TwoCell && _collectionType==DrawingsCollectionType.Worksheet)
                    {
                        string s = GetXmlNodeString("@editAs");
                        if (s == "")
                        {
                            return eEditAs.TwoCell;
                        }
                        else
                        {
                            return (eEditAs)Enum.Parse(typeof(eEditAs), s, true);
                        }
                    }
                    else
                    {
                        return CellAnchor;
                    }
                }
                catch
                {
                    return eEditAs.TwoCell;
                }
            }
            set
            {
                if (_parent != null)
                {
                    if (DrawingType == eDrawingType.Control)
                    {
                        ((ExcelControl)this).SetCellAnchor(value);
                    }
                    else if (DrawingType == eDrawingType.OleObject)
                    {
                        ((ExcelOleObject)this).SetCellAnchor(value);
                    }
                    else
                    {
                        throw (new InvalidOperationException("EditAs can't be set when a drawing is a part of a group."));
                    }
                }
                else if (CellAnchor == eEditAs.TwoCell && _collectionType==DrawingsCollectionType.Worksheet)
                {
                    string s = value.ToString();
                    SetXmlNodeString("@editAs", s.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + s.Substring(1, s.Length - 1));
                }
                else if (CellAnchor != value)
                {
                    throw (new InvalidOperationException("EditAs can only be set when CellAnchor is set to TwoCellAnchor and the drawing does not have a parent drawing."));
                }
            }
        }

        const string lockedPath = "xdr:clientData/@fLocksWithSheet";
        /// <summary>
        /// Lock drawing
        /// </summary>
        public virtual bool Locked
        {
            get
            {
                return GetXmlNodeBool(lockedPath, true);
            }
            set
            {
                SetXmlNodeBool(lockedPath, value);
            }
        }
        const string printPath = "xdr:clientData/@fPrintsWithSheet";
        /// <summary>
        /// Print drawing with sheet
        /// </summary>
        public virtual bool Print
        {
            get
            {
                return GetXmlNodeBool(printPath, true);
            }
            set
            {
                SetXmlNodeBool(printPath, value);
            }
        }
        /// <summary>
        /// Top Left position, if the shape is of the one- or two- cell anchor type
        /// Otherwise this property is set to null
        /// </summary>
        public ExcelPosition From
        {
            get;
            private set;
        }
        /// <summary>
        /// Top Left position, if the shape is of the absolute anchor type
        /// </summary>
        public ExcelDrawingCoordinate Position
        {
            get;
            internal set;
        }
        /// <summary>
        /// The extent of the shape, if the shape is of the one- or absolute- anchor type.
        /// Otherwise this property is set to null
        /// </summary>
        public ExcelDrawingSize Size
        {
            get;
            internal set;
        }
        /// <summary>
        /// Bottom right position
        /// </summary>
        public ExcelPosition To { get; private set; } = null;
        Uri _hyperLink = null;
        /// <summary>
        /// Hyperlink
        /// </summary>
        public Uri Hyperlink
        {
            get
            {
                return _hyperLink;
            }
            set
            {
                if (_hyperLink != null)
                {
                    DeleteNode(_hyperLinkPath);
                    if (HypRel != null)
                    {
                        _drawings._package.ZipPackage.DeletePart(UriHelper.ResolvePartUri(HypRel.SourceUri, HypRel.TargetUri));
                    }
                }

                if (value != null)
                {
                    if (value is ExcelHyperLink el && !string.IsNullOrEmpty(el.ReferenceAddress))
                    {
                        HypRel = _drawings.Part.CreateRelationship("#" + new ExcelAddress(el.ReferenceAddress).FullAddress, Packaging.TargetMode.Internal, ExcelPackage.schemaHyperlink);
                    }
                    else
                    {
                        HypRel = _drawings.Part.CreateRelationship(value, Packaging.TargetMode.External, ExcelPackage.schemaHyperlink);
                    }
                    SetXmlNodeString(_hyperLinkPath + "/@r:id", HypRel.Id);
                    if (Hyperlink is ExcelHyperLink excelLink)
                    {
                        SetXmlNodeString(_hyperLinkPath + "/@tooltip", excelLink.ToolTip);
                    }
                }
                _hyperLink = value;
            }
        }
        ExcelDrawingAsType _as = null;
        /// <summary>
        /// Provides access to type conversion for all top-level drawing classes.
        /// </summary>
        public ExcelDrawingAsType As
        {
            get
            {
                if (_as == null)
                {
                    _as = new ExcelDrawingAsType(this);
                }
                return _as;
            }
        }
        internal Packaging.ZipPackageRelationship HypRel { get; set; }
        /// <summary>
        /// Add new Drawing types here
        /// </summary>
        /// <param name="drawings">The drawing collection</param>
        /// <param name="node">Xml top node</param>
        /// <param name="DrawingsType">The type of collection the drawing belongs to.</param>
        /// <returns>The Drawing object</returns>
        internal static ExcelDrawing GetDrawing(ExcelDrawings drawings, XmlNode node, DrawingsCollectionType DrawingsType = DrawingsCollectionType.Worksheet)
        {
            if (node.ChildNodes.Count < 3) return null; //Invalid formatted anchor node, ignore
            XmlElement drawNode = (XmlElement)node.GetChildAtPosition(2);
            return GetDrawingFromNode(drawings, node, drawNode, null, DrawingsType);
        }

        internal static ExcelDrawing GetDrawingFromNode(ExcelDrawings drawings, XmlNode node, XmlElement drawNode, ExcelGroupShape parent = null, DrawingsCollectionType DrawingsType = DrawingsCollectionType.Worksheet)
        {
            switch (drawNode.LocalName)
            {
                case "sp":
                    return GetShapeOrControl(drawings, node, drawNode, parent, DrawingsType);
                case "pic":
                    var aPic = new ExcelPicture(drawings, node, parent, DrawingsType);
                    return aPic;
                case "graphicFrame":
                    var c= ExcelChart.GetChart(drawings, node, parent);
                    if(c!=null) //If null, the drawing is not a chart. Might be a smart art, diagram or 3d model. We return a standard drawing to retain the drawing. 
                    {
                        return c;
                    }
                    break;
                case "grpSp":
                    return new ExcelGroupShape(drawings, node, parent, DrawingsType);
                case "cxnSp":
                    return new ExcelConnectionShape(drawings, node, parent, DrawingsType);
                case "contentPart":
                    //Not handled yet, return as standard drawing below
                    break;
                case "AlternateContent":
                    XmlElement choice = drawNode.FirstChild as XmlElement;
                    if (choice != null && choice.LocalName == "Choice")
                    {
                        var req = choice.GetAttribute("Requires");  //NOTE:Can be space sparated. Might have to implement functinality for this.
                        var ns = drawNode.GetAttribute($"xmlns:{req}");
                        if (ns == "")
                        {
                            ns = choice.GetAttribute($"xmlns:{req}");
                        }
                        switch (ns)
                        {
                            case ExcelPackage.schemaChartEx2015_9_8:
                            case ExcelPackage.schemaChartEx2015_10_21:
                            case ExcelPackage.schemaChartEx2016_5_10:
                                return ExcelChart.GetChartEx(drawings, node, parent);
                            case ExcelPackage.schemaSlicer:
                                return new ExcelTableSlicer(drawings, node, parent);
                            case ExcelPackage.schemaDrawings2010:
                                if (choice.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/@uri", drawings.NameSpaceManager)?.Value == ExcelPackage.schemaSlicer2010)
                                {
                                    return new ExcelPivotTableSlicer(drawings, node, parent);
                                }
                                else if (choice.ChildNodes.Count > 0)
                                {
                                    if (choice.FirstChild.LocalName == "sp")
                                    {
                                        return GetShapeOrControl(drawings, node, (XmlElement)choice.FirstChild, parent);
                                    }
                                    else if (choice.FirstChild.LocalName == "grpSp")
                                    {
                                        return new ExcelGroupShape(drawings, choice.FirstChild, parent);
                                    }
                                }
                                break;

                        }
                    }
                    break;
            }
            return new ExcelDrawing(drawings, node, "", "");
        }

        private static ExcelDrawing GetShapeOrControl(ExcelDrawings drawings, XmlNode node, XmlElement drawNode, ExcelGroupShape parent, DrawingsCollectionType collectionType = DrawingsCollectionType.Worksheet)
        {
            var shapeId = GetControlShapeId(drawNode, drawings.NameSpaceManager, collectionType);
            var control = drawings.Worksheet.Controls.GetControlByShapeId(shapeId);
            var oleObject = control == null ? drawings.Worksheet.OleObjects.GetOleObjectByShapeId(shapeId) : null;
            if (control != null)
            {
                return ControlFactory.GetControl(drawings, drawNode, control, parent);
            }
            else if (oleObject != null)
            {
                return OleObjectFactory.GetOleObject(drawings, drawNode, oleObject, parent);
            }
            else
            {
                return new ExcelShape(drawings, node, parent, collectionType);
            }
        }

        private static int GetControlShapeId(XmlElement drawNode, XmlNamespaceManager nameSpaceManager, DrawingsCollectionType collectionType = DrawingsCollectionType.Worksheet)
        {
            var idNode = drawNode.SelectSingleNode(NamespacePrefixes[(int)collectionType] + ":nvSpPr/" + NamespacePrefixes[(int)collectionType] + ":cNvPr/@id", nameSpaceManager);
            if (idNode != null)
            {
                return int.Parse(idNode.Value);
            }
            return -1;
        }

        internal int Id
        {
            get
            {
                try
                {
                    if (_nvPrPath == "") return -1;
                    var val = GetXmlNodeInt(_nvPrPath + "/@id");
                    if (val > _drawings._nextDrawingId)
                    {
                        _drawings._nextDrawingId = val;
                    }
                    else if (val == int.MinValue)
                    {
                        val = _drawings._nextDrawingId;
                    }
                    return val;
                }
                catch
                {
                    return -1;
                }
            }
            set
            {
                try
                {
                    if (_nvPrPath == "") throw new NotImplementedException();
                    if (Id > value)
                    {
                        _drawings._nextDrawingId = Id;
                    }
                    SetXmlNodeInt(_nvPrPath + "/@id", _drawings._nextDrawingId);
                    if (this is ExcelSlicer<ExcelTableSlicerCache> ts)
                    {
                        SetXmlNodeInt(_nvPrPath + "/../../a:graphic/a:graphicData/sle:slicer/@id", value);
                        //ts._id = value;
                    }
                    else if (this is ExcelSlicer<ExcelPivotTableSlicerCache> pts)
                    {
                        SetXmlNodeInt(_nvPrPath + "/../../a:graphic/a:graphicData/sle:slicer/@id", value);
                        //pts._id = value;
                    }
                }
                catch
                {
                    throw new NotImplementedException();
                }
            }
        }
        #region "Internal sizing functions"
        internal void GetFromBounds(out int fromRow, out int fromRowOff, out int fromCol, out int fromColOff)
        {
            if (CellAnchor == eEditAs.Absolute)
            {
                GetToRowFromPixels(Position.Y, out fromRow, out fromRowOff);
                GetToColumnFromPixels(Position.X, out fromCol, out fromColOff);
            }
            else
            {
                fromRow = From.Row;
                fromRowOff = From.RowOff;
                fromCol = From.Column;
                fromColOff = From.ColumnOff;
            }
        }
        internal void GetToBounds(out int toRow, out int toRowOff, out int toCol, out int toColOff)
        {
            if (CellAnchor == eEditAs.Absolute)
            {
                GetToRowFromPixels((Position.Y + Size.Height) / EMU_PER_PIXEL, out toRow, out toRowOff);
                GetToColumnFromPixels(Position.X + Size.Width / EMU_PER_PIXEL, out toCol, out toColOff);
            }
            else
            {
                if (CellAnchor == eEditAs.TwoCell)
                {
                    toRow = To.Row;
                    toRowOff = To.RowOff;
                    toCol = To.Column;
                    toColOff = To.ColumnOff;
                }
                else
                {
                    GetToRowFromPixels(Size.Height / EMU_PER_PIXEL, out toRow, out toRowOff, From.Row, From.RowOff);
                    GetToColumnFromPixels(Size.Width / EMU_PER_PIXEL, out toCol, out toColOff, From.Column, From.ColumnOff);
                }
            }
        }
        internal int GetPixelLeft()
        {
            int pix = 0;
            if (_collectionType == DrawingsCollectionType.Chart)
            {
                if (From != null)
                {
                    return (int)(Math.Round(From.X * _drawings._screenWidth));
                }
                return 0;
            }
            if (CellAnchor == eEditAs.Absolute)
            {
                pix = Position.X / EMU_PER_PIXEL;
            }
            else
            {
                ExcelWorksheet ws = _drawings.Worksheet;
                double mdw = ws.Workbook.MaxFontWidth;

                pix = 0;
                for (int col = 0; col < From.Column; col++)
                {
                    pix += ws.GetColumnWidthPixels(col, mdw);
                }
                pix += From.ColumnOff / EMU_PER_PIXEL;
            }

            return pix;
        }
        internal int GetPixelTop()
        {
            int pix = 0;
            if (_collectionType == DrawingsCollectionType.Chart)
            {
                if (From != null)
                {
                    return (int)(Math.Round(From.Y * _drawings._screenHeight));
                }
                return 0;
            }

            if (CellAnchor == eEditAs.Absolute)
            {
                pix = Position.Y / EMU_PER_PIXEL;
            }
            else
            {
                var cache = _drawings.Worksheet.RowHeightCache;
                for (int row = 0; row < From.Row; row++)
                {
                    lock (cache)
                    {
                        if (!cache.ContainsKey(row))
                        {
                            cache.Add(row, _drawings.Worksheet.GetRowHeight(row + 1));
                        }
                    }
                    pix += (int)(cache[row] / 0.75);
                }
                pix += From.RowOff / EMU_PER_PIXEL;
            }
            return pix;
        }
        internal double GetPixelWidth()
        {
            double pix = 0;
            if (_collectionType == DrawingsCollectionType.Chart)
            {
                //var ext = (XmlElement)TopNode.SelectSingleNode("(cdr:sp|cdr:pic|cdr:cxnSp)/cdr:spPr/a:xfrm/a:ext", NameSpaceManager);
                //if (ext == null)
                //    ext = (XmlElement)TopNode.SelectSingleNode("cdr:spPr/a:xfrm/a:ext", NameSpaceManager);
                //if (ext != null)
                //{
                //    var cx = ((XmlElement)ext).GetAttribute("cx");
                //    pix = (int)double.Parse(cx);
                //}
                if(To==null)
                {
                    return Size.Width / EMU_PER_PIXEL;
                }
                return (To.X - From.X) * _drawings._screenWidth;
                //return _frmXSize.Width / EMU_PER_PIXEL;
            }
            if (CellAnchor == eEditAs.TwoCell)
            {
                ExcelWorksheet ws = _drawings.Worksheet;
                double mdw = ws.Workbook.MaxFontWidth;

                pix = -From.ColumnOff / (double)EMU_PER_PIXEL;
                for (int col = From.Column + 1; col <= To.Column; col++)
                {
                    pix += MathHelper.TruncateDouble(((256 * ws.GetColumnWidth(col) + MathHelper.TruncateDouble(128 / mdw)) / 256) * mdw);
                }

                var w = MathHelper.TruncateDouble(((256 * ws.GetColumnWidth(To.Column + 1) + MathHelper.TruncateDouble(128 / mdw)) / 256) * mdw);
                pix += Math.Min(w, Convert.ToDouble(To.ColumnOff) / EMU_PER_PIXEL);
            }
            else
            {
                pix = Size.Width / (double)EMU_PER_PIXEL;
            }
            return pix;
        }
        internal double GetPixelHeight()
        {
            double pix = 0;
            if (_collectionType == DrawingsCollectionType.Chart)
            {
                //var ext = (XmlElement)TopNode.SelectSingleNode("(cdr:sp|cdr:pic|cdr:cxnSp)/cdr:spPr/a:xfrm/a:ext", NameSpaceManager);
                //if (ext == null)
                //    ext = (XmlElement)TopNode.SelectSingleNode("cdr:spPr/a:xfrm/a:ext", NameSpaceManager);
                //if (ext != null)
                //{
                //    var cy = ((XmlElement)ext).GetAttribute("cy");
                //    pix = (int)double.Parse(cy);
                //}
                if(Size==null)
                {
                    return (To.Y - From.Y) * _drawings._screenHeight;
                }
                return Size.Height / (double)EMU_PER_PIXEL;
            }
            if (CellAnchor == eEditAs.TwoCell)
            {
                ExcelWorksheet ws = _drawings.Worksheet;

                pix = -(From.RowOff / (double)EMU_PER_PIXEL);
                for (int row = From.Row + 1; row <= To.Row; row++)
                {
                    pix += ws.GetRowHeight(row) / 0.75;
                }
                var h = ws.GetRowHeight(To.Row + 1) / 0.75;
                pix += Math.Min(h, Convert.ToDouble(To.RowOff) / EMU_PER_PIXEL);
            }
            else
            {
                pix = Size.Height / (double)EMU_PER_PIXEL;
            }
            return pix;
        }

        internal void SetPixelTop(double pixels)
        {
            _doNotAdjust = true;
            if (CellAnchor == eEditAs.Absolute)
            {
                if (_collectionType == DrawingsCollectionType.Worksheet)
                {
                    Position.Y = (int)(pixels * EMU_PER_PIXEL);
                }
                else
                {
                    From.Y= (double)pixels/_drawings._screenHeight;
                }
            }
            else
            {
                CalcRowFromPixelTop(pixels, out int row, out int rowOff);
                From.Row = row;
                From.RowOff = rowOff;
            }
            _top = pixels;
            _doNotAdjust = false;
        }

        internal void CalcRowFromPixelTop(double pixels, out int row, out int rowOff)
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            double mdw = ws.Workbook.MaxFontWidth;
            double prevPix = 0;
            double pix = ws.GetRowHeight(1) / 0.75;
            int r = 2;
            while (pix < pixels)
            {
                prevPix = pix;
                pix += (int)(ws.GetRowHeight(r++) / 0.75);
            }

            if (pix == pixels)
            {
                row = r - 1;
                rowOff = 0;
            }
            else
            {
                row = r - 2;
                rowOff = (int)(pixels - prevPix) * EMU_PER_PIXEL;
            }
        }

        internal void SetPixelLeft(double pixels)
        {
            _doNotAdjust = true;
            if (CellAnchor == eEditAs.Absolute)
            {
                if (_collectionType == DrawingsCollectionType.Worksheet)
                {
                    Position.X = (int)(pixels * EMU_PER_PIXEL);
                }
                else
                {
                    From.X = (double)pixels / _drawings._screenWidth;
                }
            }
            else
            {
                CalcColFromPixelLeft(pixels, out int col, out int colOff);
                From.Column = col;
                From.ColumnOff = colOff;
            }
            _doNotAdjust = false;

            _left = pixels;
        }
        internal void CalcColFromPixelLeft(double pixels, out int column, out int columnOff)
        {

            ExcelWorksheet ws = _drawings.Worksheet;
            double mdw = ws.Workbook.MaxFontWidth;
            double prevPix = 0;
            double pix = (int)MathHelper.TruncateDouble(((256 * ws.GetColumnWidth(1) + MathHelper.TruncateDouble(128 / mdw)) / 256) * mdw);
            int col = 2;

            while (pix < pixels)
            {
                prevPix = pix;
                pix += (int)MathHelper.TruncateDouble(((256 * ws.GetColumnWidth(col++) + MathHelper.TruncateDouble(128 / mdw)) / 256) * mdw);
            }
            if (pix == pixels)
            {
                column = col - 1;
                columnOff = 0;
            }
            else
            {
                column = col - 2;
                columnOff = (int)(pixels - prevPix) * EMU_PER_PIXEL;
            }
        }
        internal void SetPixelHeight(double pixels)
        {
            if (_collectionType == DrawingsCollectionType.Worksheet)
            {
                if (CellAnchor == eEditAs.TwoCell)
                {
                    _doNotAdjust = true;
                    GetToRowFromPixels(pixels, out int toRow, out int pixOff);
                    To.Row = toRow;
                    To.RowOff = pixOff;
                    _doNotAdjust = false;
                }
                else
                {
                    Size.Height = (long)Math.Round(pixels * EMU_PER_PIXEL);
                }
            }
            else
            {
                SetHeightChartShape(pixels);
            }
        }

        internal void GetToRowFromPixels(double pixels, out int toRow, out int rowOff, int fromRow = -1, int fromRowOff = -1)
        {
            if (fromRow < 0)
            {
                fromRow = From.Row;
                fromRowOff = From.RowOff;
            }
            ExcelWorksheet ws = _drawings.Worksheet;
            var pixOff = pixels - ((ws.GetRowHeight(fromRow + 1) / 0.75) - (fromRowOff / (double)EMU_PER_PIXEL));
            double prevPixOff = pixels;
            int row = fromRow + 1;

            while (pixOff >= 0)
            {
                prevPixOff = pixOff;
                pixOff -= (ws.GetRowHeight(++row) / 0.75);
            }
            toRow = row - 1;
            if (fromRow == toRow)
            {
                rowOff = (int)(fromRowOff + (pixels) * EMU_PER_PIXEL);
            }
            else
            {
                rowOff = (int)(prevPixOff * EMU_PER_PIXEL);
            }
        }

        internal void SetPixelWidth(double pixels)
        {
            if (_collectionType == DrawingsCollectionType.Worksheet)
            {
                if (CellAnchor == eEditAs.TwoCell)
                {
                    _doNotAdjust = true;
                    GetToColumnFromPixels(pixels, out int col, out int pixOff);

                    To.Column = col - 2;
                    To.ColumnOff = pixOff * EMU_PER_PIXEL;
                    _doNotAdjust = false;
                }
                else
                {
                    Size.Width = (int)Math.Round(pixels * EMU_PER_PIXEL);
                }
            }
            else
            {
                SetWidthChartShape((int)pixels);
            }
        }

        internal void GetToColumnFromPixels(double pixels, out int col, out int colOff, int fromColumn = -1, int fromColumnOff = -1)
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            double mdw = ws.Workbook.MaxFontWidth;
            if (fromColumn < 0)
            {
                fromColumn = From.Column;
                fromColumnOff = From.ColumnOff;
            }
            double pixOff = pixels - (MathHelper.TruncateDouble(((256 * ws.GetColumnWidth(fromColumn + 1) + MathHelper.TruncateDouble(128 / mdw)) / 256) * mdw) - fromColumnOff / EMU_PER_PIXEL);
            double offset = (double)fromColumnOff / EMU_PER_PIXEL + pixels;
            col = fromColumn + 2;
            while (pixOff >= 0)
            {
                offset = pixOff;
                pixOff -= MathHelper.TruncateDouble(((256 * ws.GetColumnWidth(col++) + MathHelper.TruncateDouble(128 / mdw)) / 256) * mdw);
            }
            colOff = (int)offset;
        }
        #endregion
        #region "Public sizing functions"
        /// <summary>
        /// Set the top left corner of a drawing. 
        /// Note that resizing columns / rows after using this function will effect the position of the drawing
        /// </summary>
        /// <param name="PixelTop">Top pixel</param>
        /// <param name="PixelLeft">Left pixel</param>
        public void SetPosition(int PixelTop, int PixelLeft)
        {
            if (_drawings._collectionType == DrawingsCollectionType.Chart)
            {
                SetPositionChartShapes(PixelTop, PixelLeft);
            }
            else
            {
                SetPosition(PixelTop, PixelLeft, true);
            }
        }
        internal void SetPosition(int PixelTop, int PixelLeft, bool adjustChildren)
        {
            _doNotAdjust = true;
            if (_width == int.MinValue)
            {
                _width = GetPixelWidth();
                _height = GetPixelHeight();
            }
            if (adjustChildren && DrawingType == eDrawingType.GroupShape)
            {
                if (_left == int.MinValue)
                {
                    _left = GetPixelLeft();
                    _top = GetPixelTop();
                }
                var grp = (ExcelGroupShape)this;
                foreach (var d in grp.Drawings)
                {
                    d.SetPosition((int)(d._top + (PixelTop - _top)), (int)(d._left + (PixelLeft - _left)));
                }
            }
            SetPixelTop(PixelTop);
            SetPixelLeft(PixelLeft);

            SetPixelWidth(_width);
            SetPixelHeight(_height);
            _doNotAdjust = false;

            if (this is ExcelOleObject ole)
            {
                ole.UpdateXml();
            }

        }


        private void SetPositionChartShapes(int PixelTop, int PixelLeft)
        {
            _top = PixelTop;
            _left = PixelLeft;
            var y = PixelTop / (_drawings._screenHeight);
            var x = PixelLeft / (_drawings._screenWidth);
            AdjustFromToXY(x, y);
            var left = (int)(From.X * _drawings._screenWidth) * EMU_PER_PIXEL;
            var top = (int)(From.Y * _drawings._screenHeight) * EMU_PER_PIXEL;
            if (_frmXPosition != null)
            {
                _frmXPosition.X = left;
                _frmXPosition.Y = top;
            }
            UpdatePositionAndSizeXml();
        }

        private void AdjustFromToXY(double x, double y)
        {
            if (y < 0)
            {
                y = 0;
            }
            else if (y > 1)
            {
                y = 1;
            }
            if (x < 0)
            {
                x = 0;
            }
            else if (x > 1)
            {
                x = 1;
            }

            if (Size==null)
            {
                var width = Math.Abs(From.X - To.X);
                var height = Math.Abs(From.Y - To.Y);

                From.X = x;
                From.Y = y;

                To.X = x + width;
                To.Y = y + height;
                if (To.X > 1)
                {
                    var diff = To.X - 1;
                    To.X = 1;
                    From.X -= diff;
                }
                if (To.Y > 1)
                {
                    var diff = To.Y - 1;
                    To.Y = 1;
                    From.Y -= diff;
                }
            }
            else
            {
                From.X = x;
                From.Y = y;
            }
        }


        /// <summary>
        /// How the drawing is anchored to the cells.
        /// This effect how the drawing will be resize
        /// <see cref="ChangeCellAnchor(eEditAs, int, int, int, int)"/>
        /// </summary>
        public eEditAs CellAnchor
        {
            get;
            protected set;
        }
        /// <summary>
        /// This will change the cell anchor type, move and resize the drawing.
        /// </summary>
        /// <param name="type">The cell anchor type to change to</param>
        /// <param name="PixelTop">The topmost pixel</param>
        /// <param name="PixelLeft">The leftmost pixel</param>
        /// <param name="width">The width in pixels</param>
        /// <param name="height">The height in pixels</param>
        public void ChangeCellAnchor(eEditAs type, int PixelTop, int PixelLeft, int width, int height)
        {
            ChangeCellAnchorTypeInternal(type);
            SetPosition(PixelTop, PixelLeft);
            SetSize(width, height);
        }
        /// <summary>
        /// This will change the cell anchor type without modifying the position and size.
        /// </summary>
        /// <param name="type">The cell anchor type to change to</param>
        public void ChangeCellAnchor(eEditAs type)
        {
            if (DrawingType == eDrawingType.Control)
            {
                throw new InvalidOperationException("Controls can't change CellAnchor. Must be TwoCell anchor. Please use EditAs property instead.");
            }
            else if (DrawingType == eDrawingType.OleObject)
            {
                throw new InvalidOperationException("Ole Objects can't change CellAnchor. Must be TwoCell anchor. Please use EditAs property instead.");
            }
            else if (_collectionType==DrawingsCollectionType.Chart && type==eEditAs.OneCell)
            {
                throw new InvalidOperationException("Drawings inside charts can't change CellAnchor to OneCell. Must be TwoCell or Absolute anchor.");
            }

            GetPositionSize();
            //Save the positions
            var top = _top;
            var left = _left;
            var width = _width;
            var height = _height;
            //Change the type
            ChangeCellAnchorTypeInternal(type);

            //Set the position and size
            SetPixelTop(top);
            SetPixelLeft(left);

            SetPixelWidth(width);
            SetPixelHeight(height);
        }
        public void GetSizeInPixels(out int width, out int height)
        {
            GetPositionSize();
            width = (int)_width;
            height = (int)_height;
        }
        private void ChangeCellAnchorTypeInternal(eEditAs type)
        {
            if (type != CellAnchor)
            {
                CellAnchor = type;
                if (_collectionType == DrawingsCollectionType.Worksheet)
                {
                    RenameNode(TopNode, "xdr", $"{type.ToEnumString()}Anchor");
                    CleanupPositionXml("xdr");
                    SetPositionProperties(_drawings, TopNode);
                    CellAnchorChanged();
                }
                else //Chart
                {
                    RenameNode(TopNode, "cdr", $"{(type==eEditAs.Absolute ? "absSizeAnchor" : "relSizeAnchor")}");
                    CleanupPositionXml("cdr");
                    SetPositionProperties(_drawings, TopNode);
                    CellAnchorChanged();
                }
            }
        }
        internal void SetCellAnchorFromNode()
        {
            if (TopNode.LocalName == "twoCellAnchor")
            {
                EditAs = CellAnchor = eEditAs.TwoCell;
            }
            else if (TopNode.LocalName == "oneCellAnchor")
            {
                CellAnchor = eEditAs.OneCell;
            }
            else
            {
                CellAnchor = eEditAs.Absolute;
            }
            SetPositionPropertiesTopDrawing(_drawings, TopNode);
        }
        internal virtual void CellAnchorChanged()
        {

        }

        private void CleanupPositionXml(string prefix)
        {
            switch (CellAnchor)
            {
                case eEditAs.OneCell:
                    DeleteNode($"{prefix}:to");
                    DeleteNode($"{prefix}:pos");
                    CreateNode($"{prefix}:from");
                    CreateNode($"{prefix}:ext");
                    break;
                case eEditAs.Absolute:
                    DeleteNode($"{prefix}:to");
                    if (_collectionType == DrawingsCollectionType.Worksheet)
                    {
                        DeleteNode($"{prefix}:from");
                        CreateNode($"{prefix}:pos");
                    }
                    else
                    {
                        CreateNode($"{prefix}:from");
                        DeleteNode($"{prefix}:to");
                    }
                    CreateNode($"{prefix}:ext");

                    break;
                default:
                    CreateNode($"{prefix}:from");
                    CreateNode($"{prefix}:to");
                    DeleteNode($"{prefix}:pos");
                    DeleteNode($"{prefix}:ext");
                    break;
            }
        }

        /// <summary>
        /// Set the top left corner of a drawing. 
        /// Note that resizing columns / rows after using this function will effect the position of the drawing
        /// </summary>
        /// <param name="Row">Start row - 0-based index.</param>
        /// <param name="RowOffsetPixels">Offset in pixels</param>
        /// <param name="Column">Start Column - 0-based index.</param>
        /// <param name="ColumnOffsetPixels">Offset in pixels</param>
        public void SetPosition(int Row, int RowOffsetPixels, int Column, int ColumnOffsetPixels)
        {
            //Throw exception if shape in Chart
            if (_collectionType == DrawingsCollectionType.Chart)
            {
                throw new InvalidOperationException("Shapes in chart does not contain row or column attributes. Use SetPosition(int PixelTop, int PixelLeft) instead.");
            }

            _doNotAdjust = true;

            if (_width == int.MinValue)
            {
                _width = GetPixelWidth();
                _height = GetPixelHeight();
            }

            From.Row = Row;
            From.RowOff = RowOffsetPixels * EMU_PER_PIXEL;
            From.Column = Column;
            From.ColumnOff = ColumnOffsetPixels * EMU_PER_PIXEL;
            if (CellAnchor == eEditAs.TwoCell)
            {
                _left = GetPixelLeft();
                _top = GetPixelTop();
            }

            SetPixelWidth(_width);
            SetPixelHeight(_height);
            _doNotAdjust = false;
            UpdatePositionAndSizeXml();
        }
        /// <summary>
        /// Set size in Percent.
        /// Note that resizing columns / rows after using this function will effect the size of the drawing
        /// </summary>
        /// <param name="Percent"></param>
        public virtual void SetSize(int Percent)
        {
            if (_drawings._collectionType == DrawingsCollectionType.Chart)
            {
                _width = Math.Round(GetPixelWidth() * ((double)Percent / 100), 0); 
                _height = Math.Round(GetPixelHeight() * ((double)Percent / 100), 0); 
                SetSizeChartShape((int)_width, (int)_height);
            }
            else
            {
                _doNotAdjust = true;
                if (_width == int.MinValue)
                {
                    _width = GetPixelWidth();
                    _height = GetPixelHeight();
                }
                _width = _width * ((double)Percent / 100);
                _height = _height * ((double)Percent / 100);

                SetPixelWidth(_width);
                SetPixelHeight(_height);
                _doNotAdjust = false;
            }
            UpdatePositionAndSizeXml();
        }


        private void SetSizeChartShape(double PixelWidth, double PixelHeight)
        {
            SetWidthChartShape(PixelWidth);
            SetHeightChartShape(PixelHeight);
        }
        private void SetWidthChartShape(double PixelWidth)
        {
            if (_frmXSize != null)
            {
                _frmXSize.Width = (long)(PixelWidth * EMU_PER_PIXEL);
            }
            if (To != null)
            {
                To.X = (From.X + PixelWidth / _drawings._screenWidth);
                if (To.X > 1) To.X = 1; else if (To.X < 0) To.X = 0;
            }
            if (Size != null)
            {
                Size.Width = (long)(PixelWidth * EMU_PER_PIXEL);
            }
        }
        private void SetHeightChartShape(double PixelHeight)
        {
            if (_frmXSize != null)
            {
                _frmXSize.Height = (long)(PixelHeight * EMU_PER_PIXEL);
            }
            if (To != null)
            {
                
                
                if (To.X > 1) To.X = 1; else if (To.X < 0) To.X = 0;

                To.Y = (From.Y + PixelHeight / _drawings._screenHeight);
                if (To.Y > 1) To.Y = 1; else if (To.Y < 0) To.Y = 0;
            }
            if (Size != null)
            {
                Size.Height = ((long)PixelHeight * EMU_PER_PIXEL);
            }
        }
        /// <summary>
        /// Set size in pixels
        /// Note that resizing columns / rows after using this function will effect the size of the drawing
        /// </summary>
        /// <param name="PixelWidth">Width in pixels</param>
        /// <param name="PixelHeight">Height in pixels</param>
        public void SetSize(int PixelWidth, int PixelHeight)
        {
            _width = PixelWidth;
            _height = PixelHeight;
            if (_drawings._collectionType == DrawingsCollectionType.Chart)
            {
                SetSizeChartShape(PixelWidth, PixelHeight);
            }
            else
            {
                _doNotAdjust = true;
                SetPixelWidth(PixelWidth);
                SetPixelHeight(PixelHeight);
                _doNotAdjust = false;
            }
            UpdatePositionAndSizeXml();
        }
        #endregion
        /// <summary>
        /// Sends the drawing to the back of any overlapping drawings.
        /// </summary>
        public void SendToBack()
        {
            _drawings.SendToBack(this);
        }
        /// <summary>
        /// Brings the drawing to the front of any overlapping drawings.
        /// </summary>
        public void BringToFront()
        {
            _drawings.BringToFront(this);
        }
        /// <summary>
        /// Group the drawing together with a list of other drawings. 
        /// <seealso cref="UnGroup(bool)"/>
        /// <seealso cref="ParentGroup"/>
        /// </summary>
        /// <param name="drawing">The drawings to group</param>
        /// <returns>The group shape</returns>
        public ExcelGroupShape Group(params ExcelDrawing[] drawing)
        {
            ExcelGroupShape grp = _parent;
            foreach (var d in drawing)
            {
                ExcelGroupShape.Validate(d, _drawings, grp);
                if (d._parent != null) grp = d._parent;
            }
            if (grp == null)
            {
                grp = _drawings.AddGroupDrawing(_drawings._collectionType);
            }

            grp.Drawings.AddDrawing(this);

            foreach (var d in drawing)
            {
                grp.Drawings.AddDrawing(d);
            }

            grp.SetPositionAndSizeFromChildren();
            return grp;
        }
        internal XmlElement GetXfrmNode(XmlNode node)
        {
            if (node == null) return null;
            if (node.LocalName == "AlternateContent")
            {
                node = node.FirstChild.FirstChild;
            }

            if (node.LocalName == "sp" || node.LocalName == "pic" || node.LocalName == "cxnSp")
            {
                return (XmlElement)CreateNode(node, NamespacePrefixes[_prefixIndex] + ":spPr/a:xfrm");
            }
            else if (node.LocalName == "grpSp")
            {
                return (XmlElement)CreateNode(node, NamespacePrefixes[_prefixIndex] + ":grpSpPr/a:xfrm");
            }
            else if (node.LocalName == "graphicFrame")
            {
                return (XmlElement)CreateNode(node, NamespacePrefixes[_prefixIndex] + ":xfrm");
            }
            return null;
        }

        /// <summary>
        /// Will ungroup this drawing or the entire group, if this drawing is grouped together with other drawings.
        /// If this drawings is not grouped an InvalidOperationException will be returned.
        /// </summary>
        /// <param name="ungroupThisItemOnly">If true this drawing will be removed from the group. 
        /// If it is false, the whole group will be disbanded. If true only this drawing will be removed.
        /// </param>
        public void UnGroup(bool ungroupThisItemOnly = true)
        {
            if (_parent == null)
            {
                throw new InvalidOperationException("Cannot ungroup this drawing. This drawing is not part of a group");
            }
            var prevParent = _parent;
            if (ungroupThisItemOnly)
            {
                _parent.Drawings.Remove(this);
            }
            else
            {
                _parent.Drawings.Clear();
            }
            if (prevParent.Drawings.Count <= 0)
            {
                prevParent.DeleteMe();
            }
        }
        /// <summary>
        /// If the drawing is grouped this property contains the Group drawing containing the group.
        /// Otherwise this property is null
        /// </summary>
        public ExcelGroupShape ParentGroup
        {
            get
            {
                return _parent;
            }
        }

        internal ExcelDrawingCustomGeometry CustomGeom { get; private set; }

        internal virtual void DeleteMe()
        {
            TopNode.ParentNode.RemoveChild(TopNode);
        }

        /// <summary>
        /// Dispose the object
        /// </summary>
        public virtual void Dispose()
        {
            //TopNode = null;
        }
        internal void GetPositionSize()
        {
            if (_doNotAdjust) return;
            _top = GetPixelTop();
            _left = GetPixelLeft();
            _height = GetPixelHeight();
            _width = GetPixelWidth();
        }
        /// <summary>
        /// Will adjust the position and size of the drawing according to changes in font of rows and to the Normal style.
        /// This method will be called before save, so use it only if you need the coordinates of the drawing.
        /// </summary>
        public void AdjustPositionAndSize()
        {
            if (_drawings.Worksheet.Workbook._package.DoAdjustDrawings == false) return;
            _drawings.Worksheet.Workbook._package.DoAdjustDrawings = false;
            if (EditAs == eEditAs.Absolute)
            {
                SetPixelLeft(_left);
                SetPixelTop(_top);
            }
            if (EditAs == eEditAs.Absolute || EditAs == eEditAs.OneCell)
            {
                SetPixelHeight(_height);
                SetPixelWidth(_width);
            }

            _drawings.Worksheet.Workbook._package.DoAdjustDrawings = true;
        }
        internal void UpdatePositionAndSizeXml()
        {
            From?.UpdateXml();
            To?.UpdateXml();
            Size?.UpdateXml();
            Position?.UpdateXml();
            if(_collectionType==DrawingsCollectionType.Chart)
            {
                _frmXPosition?.UpdateXml();
                _frmXSize?.UpdateXml();
            }
        }


        internal XmlElement CreateShapeNode()
        {
            XmlElement shapeNode;
            switch (_drawings._collectionType)
            {
                case DrawingsCollectionType.Chart:
                    shapeNode = TopNode.OwnerDocument.CreateElement("cdr", "sp", ExcelPackage.schemaChartDrawing);
                    break;
                case DrawingsCollectionType.Worksheet:
                default:
                    shapeNode = TopNode.OwnerDocument.CreateElement("xdr", "sp", ExcelPackage.schemaSheetDrawings);
                    break;
            }
            shapeNode.SetAttribute("macro", "");
            shapeNode.SetAttribute("textlink", "");
            TopNode.AppendChild(shapeNode);
            return shapeNode;
        }
        internal XmlElement CreateClientData(bool printsWithSheet = true)
        {
            XmlElement clientDataNode = TopNode.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings);
            if (printsWithSheet)
            {
                clientDataNode.SetAttribute("fPrintsWithSheet", "0");
            }
            var parentNode = TopNode.GetChildAtPosition(2).GetChildAtPosition(0).GetChildAtPosition(0);
            parentNode.AppendChild(clientDataNode);
            //InserAfter(top)
            TopNode.AppendChild(clientDataNode);
            return clientDataNode;
        }

        /// <summary>
        /// Create a copy of target chart.
        /// </summary>
        /// <param name="targetChart"></param>
        /// <param name="PixelTop"></param>
        /// <param name="PixelLeft"></param>
        /// <exception cref="NotSupportedException"></exception>
        public void Copy(ExcelChartStandard targetChart, int PixelTop = -1, int PixelLeft = -1)
        {
            if (this is ExcelShape || this is ExcelPicture || this is ExcelGroupShape)
            {
                XmlNode drawNode = null;
                switch (DrawingType)
                {
                    case eDrawingType.Shape:
                        drawNode = CopyShape(targetChart);
                        break;
                    case eDrawingType.Picture:
                        drawNode = CopyPicture(targetChart);
                        break;
                    case eDrawingType.GroupShape:
                        drawNode = CopyGroupShape(targetChart);
                        break;
                }
                if (targetChart is ExcelChartStandard chartStandard)
                {
                    var copy = GetDrawing(chartStandard.Drawings._drawings, drawNode, DrawingsCollectionType.Chart);
                    chartStandard.Drawings.AddDrawingInternal(copy);
                    if (PixelTop > 0 || PixelLeft > 0)
                    {
                        copy.SetPosition(PixelTop, PixelLeft);
                    }
                }
            }
            else
            {
                throw new NotSupportedException("Charts only supports shapes, pictures and group shapes containing only shapes or pictures.");
            }
        }


        /// <summary>
        /// Copies the drawing to the supplied worksheets. The copy will be positioned using the <paramref name="row"/> and <paramref name="col"/> parameters
        /// </summary>
        /// <param name="worksheet">The worksheet where the drawing will be placed.</param>
        /// <param name="row">The top row where the drawing will be placed.</param>
        /// <param name="col">The left column where the drawing will be placed.</param>
        /// <param name="rowOffset">Row offset in pixels from the row start positions. int.MinValue </param>
        /// <param name="colOffset">Column offset in pixels from the column start position</param>
        public ExcelDrawing Copy(ExcelWorksheet worksheet, int row, int col, int rowOffset = int.MinValue, int colOffset = int.MinValue)
        {
            XmlNode drawNode = null;
            if (From == null)
            {
                if (rowOffset == int.MinValue || colOffset == int.MinValue)
                {
                    GetFromBounds(out _, out int ro, out _, out int co);
                    if (rowOffset == int.MinValue)
                    {
                        rowOffset = ro;
                    }
                    if (colOffset == int.MinValue)
                    {
                        colOffset = co;
                    }
                }
            }
            else
            {
                if (rowOffset == int.MinValue)
                {
                    rowOffset = From.RowOff / 9525;
                }
                if (colOffset == int.MinValue)
                {
                    colOffset = From.ColumnOff / 9525;
                }
            }

            switch (DrawingType)
            {
                case eDrawingType.Shape:
                    drawNode = CopyShape(worksheet);
                    break;
                case eDrawingType.Picture:
                    drawNode = CopyPicture(worksheet);
                    break;
                case eDrawingType.Chart:
                    drawNode = CopyChart(worksheet);
                    break;
                case eDrawingType.Slicer:
                    drawNode = CopySlicer(worksheet);
                    break;
                case eDrawingType.Control:
                    drawNode = CopyControl(worksheet, row, col, rowOffset, colOffset);
                    break;
                case eDrawingType.GroupShape:
                    drawNode = CopyGroupShape(worksheet);
                    break;
                case eDrawingType.OleObject:
                    drawNode = CopyOleObject(worksheet, row, col, rowOffset, colOffset);
                    return GetDrawing(worksheet._drawings, drawNode);
            }
            //Set position of the drawing copy.
            var copy = GetDrawing(worksheet._drawings, drawNode);
            worksheet.Drawings.AddDrawingInternal(copy);
            var width = GetPixelWidth();
            var height = GetPixelHeight();
            copy.SetPosition(row, rowOffset, col, colOffset);
            copy.SetPixelWidth(width);
            copy.SetPixelHeight(height);
            copy.GetPositionSize();
            return copy;
        }

        private XmlNode CopyGroupShape(ExcelChartStandard targetChart)
        {
            var drawNode = targetChart.Drawings.CreateDocumentAndTopNodeChartDrawings(targetChart);
            drawNode.InnerXml = TopNode.InnerXml;
            CopyGroupShape(targetChart, this, drawNode.ChildNodes[2]);
            return drawNode;
        }

        private void CopyGroupShape(ExcelChartStandard targetChart, ExcelDrawing sourceDrawing, XmlNode targetDrawNode, ExcelGroupShape parent = null)
        {
            if (sourceDrawing is ExcelShape shape)
            {
                sourceDrawing.CopyShape(targetChart, true, targetDrawNode);
            }
            else if (sourceDrawing is ExcelPicture picture)
            {
                sourceDrawing.CopyPicture(targetChart, true, targetDrawNode);
            }
            else if (sourceDrawing is ExcelGroupShape groupShape)
            {
                int nodeIndex = 2;
                for (int j = 0; j < groupShape.Drawings.Count; j++)
                {
                    //Start at index 2 but child nodes must be incremented by 1 each loop so that we check the next node.
                    CopyGroupShape(targetChart, groupShape.Drawings[j], targetDrawNode.ChildNodes[nodeIndex++], groupShape);
                }
            }
            else
            {
                throw new NotSupportedException("Charts only supports shapes, pictures and group shapes containing only shapes or pictures.");
            }
        }

        private XmlNode CopyGroupShape(ExcelWorksheet worksheet)
        {
            //Create node in drawing.xml
            var drawNode = worksheet.Drawings.CreateDocumentAndTopNode(CellAnchor, false);
            drawNode.InnerXml = TopNode.InnerXml;
            CopyGroupShape(worksheet, this, drawNode.ChildNodes[2]);
            return drawNode;
        }

        private void CopyGroupShape(ExcelWorksheet targetWorksheet, ExcelDrawing sourceDrawing, XmlNode targetDrawNode, ExcelGroupShape parent = null)
        {
            if (sourceDrawing is ExcelChart chart)
            {
                sourceDrawing.CopyChart(targetWorksheet, true, targetDrawNode);
            }
            if (sourceDrawing is ExcelPicture pic)
            {
                sourceDrawing.CopyPicture(targetWorksheet, true, targetDrawNode);
            }
            if (sourceDrawing is ExcelControl ctrl)
            {
                sourceDrawing.CopyControl(targetWorksheet, 0, 0, 0, 0, true, targetDrawNode);
            }
            else if (sourceDrawing is ExcelShape shape)
            {
                sourceDrawing.CopyShape(targetWorksheet, true, targetDrawNode);
            }
            else if (sourceDrawing is ExcelTableSlicer tSlicer)
            {
                sourceDrawing.CopySlicer(targetWorksheet, true, targetDrawNode);
            }
            else if (sourceDrawing is ExcelPivotTableSlicer ptSlicer)
            {
                sourceDrawing.CopySlicer(targetWorksheet, true, targetDrawNode);
            }
            else if (sourceDrawing is ExcelOleObject ole)
            {
                sourceDrawing.CopyOleObject(targetWorksheet, 0, 0, 0, 0, true, targetDrawNode, parent);
            }
            else if (sourceDrawing is ExcelGroupShape groupShape)
            {
                int nodeIndex = 2;
                for (int j = 0; j < groupShape.Drawings.Count; j++)
                {
                    //Start at index 2 but child nodes must be incremented by 1 each loop so that we check the next node.
                    CopyGroupShape(targetWorksheet, groupShape.Drawings[j], targetDrawNode.ChildNodes[nodeIndex++], groupShape);
                }
            }
        }

        private XmlNode CopySlicer(ExcelWorksheet worksheet, bool isGroupShape = false, XmlNode groupDrawNode = null)
        {
            //can't copy to another workbook unless we also copy the table. (Need to check for table somehow...)
            if (worksheet.Workbook != _drawings.Worksheet.Workbook)
            {
                throw new InvalidOperationException("Table slicers can't be copied from one workbook to another.");
            }

            //Create node in drawing.xml
            XmlNode drawNode = null;
            if (isGroupShape)
            {
                drawNode = groupDrawNode;
            }
            else
            {
                drawNode = worksheet.Drawings.CreateDocumentAndTopNode(CellAnchor, false);
                drawNode.InnerXml = TopNode.InnerXml;
            }

            //Create copy of source worksheet node in target worksheet.xml
            XmlNode wsSlicerNode = worksheet.TopNode.SelectSingleNode("d:extLst/d:ext/x14:slicerList/x14:slicer", worksheet.NameSpaceManager);
            if (worksheet != _drawings.Worksheet)
            {
                if (wsSlicerNode == null)
                {
                    ((XmlElement)worksheet.TopNode).SetAttribute("xmlns:x14", ExcelPackage.schemaMainX14);   //Make sure the namespace exists
                    var slicerNode = worksheet.CreateNode("d:extLst");
                    slicerNode.InnerXml = _drawings.Worksheet.TopNode.SelectSingleNode("d:extLst", _drawings.Worksheet.NameSpaceManager).InnerXml;
                }
            }

            ////Set Name in drawingXML
            var drawNodeName = drawNode.SelectSingleNode("mc:AlternateContent/mc:Choice/xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr", worksheet._drawings.NameSpaceManager);
            if (drawNodeName == null && isGroupShape)
            {
                drawNodeName = drawNode.SelectSingleNode("mc:Choice/xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr", worksheet._drawings.NameSpaceManager);
            }
            var slicerName = worksheet.Workbook.GetSlicerName(drawNodeName.Attributes["name"].Value); //worksheet._drawings.GetUniqueDrawingName(drawNodeName.Attributes["name"].Value) + "fsgs";
            drawNodeName.Attributes["name"].Value = slicerName;
            var drawNodeSlicerName = drawNode.SelectSingleNode("mc:AlternateContent/mc:Choice/xdr:graphicFrame/a:graphic/a:graphicData/sle:slicer", worksheet._drawings.NameSpaceManager);
            if (drawNodeSlicerName == null && isGroupShape)
            {
                drawNodeSlicerName = drawNode.SelectSingleNode("mc:Choice/xdr:graphicFrame/a:graphic/a:graphicData/sle:slicer", worksheet._drawings.NameSpaceManager);
            }
            drawNodeSlicerName.Attributes["name"].Value = slicerName;

            //Copy Slicer xml node
            Uri uri;
            ZipPackagePart part = null;
            ZipPackageRelationship relationship = null;
            bool isNewPart = false;
            if (wsSlicerNode == null)
            {
                var id = worksheet.SheetId;
                uri = XmlHelper.GetNewUri(worksheet.Part.Package, "/xl/slicers/slicer{0}.xml", ref id);
                part = worksheet.Part.Package.CreatePart(uri, "application/vnd.ms-excel.slicer+xml", worksheet.Part.Package.Compression);
                relationship = worksheet.Part.CreateRelationship(uri, Packaging.TargetMode.Internal, ExcelPackage.schemaRelationshipsSlicer);
                isNewPart = true;
            }
            else
            {
                part = worksheet.SlicerXmlSources._part;
            }

            var xmlTarget = new XmlDocument();
            ExcelSlicerXmlSource xmlSource = null;
            string name = string.Empty;
            if (this is ExcelTableSlicer ets)
            {
                xmlSource = _drawings.Worksheet.SlicerXmlSources._list.Find(x => x == ets._xmlSource);
                name = ets.Name;
            }
            else if (this is ExcelPivotTableSlicer epts)
            {
                xmlSource = _drawings.Worksheet.SlicerXmlSources._list.Find(x => x == epts._xmlSource);
                name = epts.Name;
            }
            //If different drawings create a new xml. (Maybe check for exsisting xml in new drawings and append instead)
            if (_drawings != worksheet._drawings)
            {
                if (isNewPart)
                {
                    XmlHelper.LoadXmlSafe(xmlTarget, "<slicers xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" mc:Ignorable=\"x xr10\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"/>", Encoding.UTF8);
                }
                else
                {
                    xmlTarget = worksheet.SlicerXmlSources._list.Find(x => x.Type == eSlicerSourceType.Table).XmlDocument; //hï¿½ller en kopi, need to skriv ref...
                }
            }
            else
            {
                xmlTarget = xmlSource.XmlDocument;
            }

            //Set name in SlicerXML
            var slicerNodes = xmlSource.XmlDocument.LastChild.ChildNodes;
            XmlNode importNode = null;
            foreach (XmlNode node in slicerNodes)
            {
                if (node.Attributes["name"].Value == name)
                {
                    importNode = node.CloneNode(true);
                    break;
                }
            }
            importNode.Attributes["name"].Value = slicerName;
            var newNode = xmlTarget.ImportNode(importNode, true);
            xmlTarget.LastChild.AppendChild(newNode);
            var stream = new StreamWriter(part.GetStream(FileMode.OpenOrCreate, FileAccess.Write));
            xmlTarget.Save(stream);

            if (isNewPart)
            {
                //Now create the new relationship between the worksheet and the slicer.
                var relNode = (XmlElement)(worksheet.WorksheetXml.DocumentElement.SelectSingleNode($"d:extLst/d:ext/x14:slicerList/x14:slicer[@r:id='{xmlSource.Rel.Id}']", worksheet.NameSpaceManager));
                relNode.Attributes["r:id"].Value = relationship.Id;
            }
            return drawNode;
        }

        private XmlNode CopyControl(ExcelWorksheet worksheet, int row, int col, int rowOffset, int colOffset, bool isGroupShape = false, XmlNode groupDrawNode = null)
        {
            XmlNode drawNode = null;
            if (isGroupShape && groupDrawNode != null)
            {
                drawNode = groupDrawNode.FirstChild;
            }
            else
            {
                //Create node in drawing.xml
                drawNode = worksheet.Drawings.CreateDocumentAndTopNode(CellAnchor, true);
                drawNode.InnerXml = TopNode.InnerXml;
            }
            //Update DrawNode Id
            var controlId = (++worksheet._nextControlId).ToString();
            var drawIdNode = drawNode.SelectSingleNode("xdr:sp/xdr:nvSpPr/xdr:cNvPr", worksheet.NameSpaceManager);
            drawIdNode.Attributes["id"].Value = controlId;
            var drawSpIdNode = drawIdNode.SelectSingleNode("a:extLst/a:ext/a14:compatExt", _drawings.NameSpaceManager);
            var spid = drawSpIdNode.Attributes["spid"].Value = "_x0000_s" + controlId;

            //Create worksheet node
            var control = this as ExcelControl;
            XmlNode controlNode = worksheet.CreateControlContainerNode();
            ((XmlElement)worksheet.TopNode).SetAttribute("xmlns:xdr", ExcelPackage.schemaSheetDrawings);   //Make sure the namespace exists
            ((XmlElement)worksheet.TopNode).SetAttribute("xmlns:x14", ExcelPackage.schemaMainX14);   //Make sure the namespace exists
            ((XmlElement)worksheet.TopNode).SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);   //Make sure the namespace exists
            controlNode.InnerXml = control._control.TopNode.ParentNode.InnerXml;
            controlNode.FirstChild.Attributes["shapeId"].Value = controlId;
            WorksheetCopyHelper.CopyControl(worksheet._package, worksheet, control);

            //Create vml
            worksheet.VmlDrawings.AddControl(control, spid);
            var vmlId = worksheet.VmlDrawings._drawings[worksheet.VmlDrawings._drawings.Count - 1].TopNode;
            vmlId.Attributes["spid"].Value = spid;
            if (!isGroupShape)
            {
                //Create the copy
                var copy = GetDrawing(worksheet._drawings, drawNode);
                copy.EditAs = ExcelControl.GetControlEditAs(control.ControlType);
                var width = GetPixelWidth();
                var height = GetPixelHeight();
                copy.SetPosition(row, rowOffset, col, colOffset);
                copy.SetPixelWidth(width);
                copy.SetPixelHeight(height);
                copy.GetPositionSize();

                //Update position in worksheet xml
                var fromCol = controlNode.SelectSingleNode("d:control/d:controlPr/d:anchor/d:from/xdr:col", worksheet.NameSpaceManager);
                var fromColOff = controlNode.SelectSingleNode("d:control/d:controlPr/d:anchor/d:from/xdr:colOff", worksheet.NameSpaceManager);
                var fromRow = controlNode.SelectSingleNode("d:control/d:controlPr/d:anchor/d:from/xdr:row", worksheet.NameSpaceManager);
                var fromRowOff = controlNode.SelectSingleNode("d:control/d:controlPr/d:anchor/d:from/xdr:rowOff", worksheet.NameSpaceManager);
                fromCol.InnerText = copy.From.Column.ToString();
                fromColOff.InnerText = copy.From.ColumnOff.ToString();
                fromRow.InnerText = copy.From.Row.ToString();
                fromRowOff.InnerText = copy.From.RowOff.ToString();
                var toCol = controlNode.SelectSingleNode("d:control/d:controlPr/d:anchor/d:to/xdr:col", worksheet.NameSpaceManager);
                var toColOff = controlNode.SelectSingleNode("d:control/d:controlPr/d:anchor/d:to/xdr:colOff", worksheet.NameSpaceManager);
                var toRow = controlNode.SelectSingleNode("d:control/d:controlPr/d:anchor/d:to/xdr:row", worksheet.NameSpaceManager);
                var toRowOff = controlNode.SelectSingleNode("d:control/d:controlPr/d:anchor/d:to/xdr:rowOff", worksheet.NameSpaceManager);
                toCol.InnerText = copy.To.Column.ToString();
                toColOff.InnerText = copy.To.ColumnOff.ToString();
                toRow.InnerText = copy.To.Row.ToString();
                toRowOff.InnerText = copy.To.RowOff.ToString();

                //Update position in drawing vml
                var vmlPosition = vmlId.SelectSingleNode("x:ClientData/x:Anchor", worksheet._vmlDrawings.NameSpaceManager);
                vmlPosition.InnerXml = copy.From.Column + ", " + copy.From.ColumnOff + ", " + copy.From.Row + ", " + copy.From.RowOff + ", " +
                                        copy.To.Column + ", " + copy.To.ColumnOff + ", " + copy.To.Row + ", " + copy.To.RowOff;
            }
            return drawNode;
        }

        private XmlNode CopyOleObject(ExcelWorksheet worksheet, int row, int col, int rowOffset, int colOffset, bool isGroupShape = false, XmlNode groupDrawNode = null, ExcelGroupShape parent = null)
        {
            var ole = this as ExcelOleObject;
            //copy drawing
            XmlNode drawNode = null;
            if (isGroupShape && groupDrawNode != null)
            {
                drawNode = groupDrawNode.FirstChild;
            }
            else
            {
                //Create node in drawing.xml
                drawNode = worksheet.Drawings.CreateDocumentAndTopNode(CellAnchor, true);
                drawNode.InnerXml = TopNode.InnerXml;
            }

            //create worksheet node
            XmlNode oleNode = worksheet.CreateOleContainerNode();
            ((XmlElement)worksheet.TopNode).SetAttribute("xmlns:xdr", ExcelPackage.schemaSheetDrawings);   //Make sure the namespace exists
            ((XmlElement)worksheet.TopNode).SetAttribute("xmlns:x14", ExcelPackage.schemaMainX14);   //Make sure the namespace exists
            ((XmlElement)worksheet.TopNode).SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);   //Make sure the namespace exists
            XmlNode newNode;
            if (ole._oleObject.TopNode.OwnerDocument == oleNode.OwnerDocument)
            {
                newNode = ole._oleObject.TopNode.ParentNode.ParentNode.CloneNode(true);
            }
            else
            {
                newNode = oleNode.OwnerDocument.ImportNode(ole._oleObject.TopNode.ParentNode.ParentNode, true);
            }
            oleNode.AppendChild(newNode);
            //Copy OleObject & Image
            var shapeId = WorksheetCopyHelper.CopyOleObject(worksheet._package, worksheet, ole, worksheet._drawings.DrawingXml);

            if (!isGroupShape)
            {
                //Create the copy
                var copy = GetDrawing(worksheet._drawings, drawNode);
                var width = GetPixelWidth();
                var height = GetPixelHeight();
                copy.SetPosition(row, rowOffset, col, colOffset);
                copy.SetPixelWidth(width);
                copy.SetPixelHeight(height);
                copy.GetPositionSize();

                //Update position in worksheet xml
                var fromCol = newNode.SelectSingleNode("mc:Choice/d:oleObject/d:objectPr/d:anchor/d:from/xdr:col", worksheet.NameSpaceManager);
                var fromColOff = newNode.SelectSingleNode("mc:Choice/d:oleObject/d:objectPr/d:anchor/d:from/xdr:colOff", worksheet.NameSpaceManager);
                var fromRow = newNode.SelectSingleNode("mc:Choice/d:oleObject/d:objectPr/d:anchor/d:from/xdr:row", worksheet.NameSpaceManager);
                var fromRowOff = newNode.SelectSingleNode("mc:Choice/d:oleObject/d:objectPr/d:anchor/d:from/xdr:rowOff", worksheet.NameSpaceManager);
                fromCol.InnerText = copy.From.Column.ToString();
                fromColOff.InnerText = copy.From.ColumnOff.ToString();
                fromRow.InnerText = copy.From.Row.ToString();
                fromRowOff.InnerText = copy.From.RowOff.ToString();
                var toCol = newNode.SelectSingleNode("mc:Choice/d:oleObject/d:objectPr/d:anchor/d:to/xdr:col", worksheet.NameSpaceManager);
                var toColOff = newNode.SelectSingleNode("mc:Choice/d:oleObject/d:objectPr/d:anchor/d:to/xdr:colOff", worksheet.NameSpaceManager);
                var toRow = newNode.SelectSingleNode("mc:Choice/d:oleObject/d:objectPr/d:anchor/d:to/xdr:row", worksheet.NameSpaceManager);
                var toRowOff = newNode.SelectSingleNode("mc:Choice/d:oleObject/d:objectPr/d:anchor/d:to/xdr:rowOff", worksheet.NameSpaceManager);
                toCol.InnerText = copy.To.Column.ToString();
                toColOff.InnerText = copy.To.ColumnOff.ToString();
                toRow.InnerText = copy.To.Row.ToString();
                toRowOff.InnerText = copy.To.RowOff.ToString();
                copy.From.UpdateXml();
                copy.To.UpdateXml();
            }
            var oleInternal = new OleObjectInternal(worksheet.NameSpaceManager, newNode.FirstChild.FirstChild);
            int shapeIdKey = int.Parse(shapeId);
            if (!worksheet.OleObjects._dict.ContainsKey(shapeIdKey))
                worksheet.OleObjects._dict.Add(shapeIdKey, oleInternal);
            var oleObject = OleObjectFactory.GetOleObject(worksheet.Drawings, drawNode.SelectSingleNode("xdr:sp", NameSpaceManager) as XmlElement, oleInternal, parent) as ExcelOleObject;
            oleObject.Name = worksheet.Drawings.GetUniqueDrawingName(oleObject.Name);
            worksheet.Drawings.AddDrawingInternal(oleObject);
            return drawNode;
        }

        private XmlNode CopyChart(ExcelWorksheet worksheet, bool isGroupShape = false, XmlNode groupDrawNode = null)
        {
            XmlNode drawNode = null;
            ExcelChart targetChart = null;
            var origialChart = this as ExcelChart;
            if (isGroupShape && groupDrawNode != null)
            {
                drawNode = groupDrawNode;
            }
            else
            {
                //Create node in drawing.xml
                drawNode = worksheet.Drawings.CreateDocumentAndTopNode(CellAnchor, false);
                drawNode.InnerXml = TopNode.InnerXml;
            }
            //get relationship node in drawing.xml
            var relNode = drawNode.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart/@r:id", NameSpaceManager);
            if (relNode == null)
            {
                relNode = drawNode.SelectSingleNode("a:graphic/a:graphicData/c:chart/@r:id", NameSpaceManager);
            }
            if (relNode != null && _drawings.Part.RelationshipExists(relNode.Value))
            {
                WorksheetCopyHelper.CopyChartRelations(origialChart, worksheet, worksheet._drawings.Part, worksheet._drawings.DrawingXml, _drawings.Worksheet, drawNode);
                //Update the copied charts id and name
                if (isGroupShape)
                {
                    var chartAttr = groupDrawNode.SelectSingleNode("xdr:nvGraphicFramePr/xdr:cNvPr", worksheet._drawings.NameSpaceManager);
                    chartAttr.Attributes["name"].Value = worksheet._drawings.GetUniqueDrawingName(origialChart.Name);
                    chartAttr.Attributes["id"].Value = (_drawings._nextDrawingId++).ToString();
                }
                else
                {
                    targetChart = ExcelChart.GetChart(worksheet.Drawings, drawNode);
                    targetChart.Name = worksheet._drawings.GetUniqueDrawingName(origialChart.Name);
                    targetChart.Id = _drawings._nextDrawingId++;
                }
            }
            return drawNode;
        }

        private XmlNode CopyPicture(ExcelChartStandard targetChart, bool isGroupShape = false, XmlNode groupDrawNode = null)
        {
            XmlNode drawNode = null;
            if (isGroupShape && groupDrawNode != null)
            {
                drawNode = groupDrawNode;
                groupDrawNode.SelectSingleNode("cdr:nvPicPr/cdr:cNvPr", targetChart.Drawings.NameSpaceManager).Attributes["id"].Value = (++targetChart.Drawings._nextDrawingId).ToString();
            }
            else
            {
                drawNode = targetChart.Drawings.CreateDrawingXmlChartDrawings(targetChart);
                drawNode.InnerXml = TopNode.InnerXml;
            }
            if (targetChart.Drawings._drawings != _drawings)
            {
                var relNode = drawNode.SelectSingleNode("cdr:pic/cdr:blipFill/a:blip/@r:embed", NameSpaceManager);
                if (relNode != null && _drawings.Part.RelationshipExists(relNode.Value))
                {
                    var rel = _drawings.Part.GetRelationship(relNode.Value);
                    //Create new relation id if no relation exsist or if it's a different worksheet. Otherwise asign the exsisting relationship Id
                    var newRel = targetChart.Drawings.Part.CreateRelationshipFromCopy(rel);
                    relNode.Value = newRel.Id;
                }
            }
            if (!isGroupShape)
            {
                var targetPic = GetDrawing(targetChart.Drawings._drawings, drawNode, DrawingsCollectionType.Chart) as ExcelPicture;
                targetPic.Id = ++targetChart.Drawings._nextDrawingId;
                targetPic.Name = targetChart._drawings.GetUniqueDrawingName(this.Name);
            }
            return drawNode;
        }

        private XmlNode CopyPicture(ExcelWorksheet targetWorksheet, bool isGroupShape = false, XmlNode groupDrawNode = null)
        {
            XmlNode drawNode = null;

            var targetWorkbook = targetWorksheet.Workbook;
            var targetPackage = targetWorkbook._package;

            if (isGroupShape && groupDrawNode != null)
            {
                drawNode = groupDrawNode;
                groupDrawNode.SelectSingleNode("xdr:nvPicPr/xdr:cNvPr", targetWorksheet._drawings.NameSpaceManager).Attributes["id"].Value = (++targetWorksheet.Drawings._nextDrawingId).ToString();
            }
            else
            {
                //Create node in drawing.xml
                drawNode = targetWorksheet.Drawings.CreateDocumentAndTopNode(CellAnchor, false);
                drawNode.InnerXml = TopNode.InnerXml;
            }
            //If same drawings object, we are done.
            if (targetWorksheet._drawings != _drawings)
            {
                //Get the relation node
                var relNode = drawNode.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager);
                if (relNode == null)
                {
                    relNode = drawNode.SelectSingleNode("xdr:blipFill/a:blip/@r:embed", NameSpaceManager);
                }

                if (relNode != null && _drawings.Part.RelationshipExists(relNode.Value))
                {
                    var srcsRel = _drawings.Part.GetRelationship(relNode.Value);
                    ZipPackageRelationship newRel = null;

                    //Copy image file to new workbook if target worksheet is in a different workbook.
                    if (targetWorkbook != _drawings.Worksheet.Workbook)
                    {
                        var uri = UriHelper.ResolvePartUri(srcsRel.SourceUri, srcsRel.TargetUri);
                        var imagePart = _drawings.Worksheet.Workbook._package.ZipPackage.GetPart(uri);

                        var imageStream = (MemoryStream)imagePart.GetStream(FileMode.Open, FileAccess.Read);
                        var image = new byte[imageStream.Length];

                        imageStream.Seek(0, SeekOrigin.Begin);
                        imageStream.Read(image, 0, (int)imageStream.Length);

                        var imageInfo = targetPackage.PictureStore.GetImageInfo(image);

                        if (imageInfo == null)
                        {
                            var info = new FileInfo(uri.OriginalString);
                            Uri absUri = GetNewUri(targetPackage.ZipPackage, "/xl/media/image{0}" + info.Extension);

                            newRel = targetWorksheet._drawings.Part.CreateRelationshipFromCopy(srcsRel);

                            var relativeUri = UriHelper.GetRelativeUri(newRel.SourceUri, absUri);
                            newRel.TargetUri = relativeUri;

                            var copyPart = targetPackage.ZipPackage.CreatePart(absUri, imagePart.ContentType);
                            var copyStream = (MemoryStream)copyPart.GetStream(FileMode.Create, FileAccess.Write);
                            copyStream.Write(image, 0, image.Length);

                            relNode.Value = newRel.Id;
                        }
                        else
                        {
                            var relativeUri = UriHelper.GetRelativeUri(srcsRel.SourceUri, imageInfo.Uri);
                            var exisistingRel = targetWorksheet._drawings.Part.GetRelationshipsByType(srcsRel.RelationshipType).Where(x => x.TargetUri == relativeUri).FirstOrDefault();
                            //Create new relation id if no relation exists. Otherwise asign the existing relationship Id
                            if (exisistingRel == null)
                            {
                                newRel = targetWorksheet._drawings.Part.CreateRelationshipFromCopy(srcsRel);
                                relNode.Value = newRel.Id;
                            }
                            else
                            {
                                relNode.Value = exisistingRel.Id;
                            }
                        }
                    }
                    else
                    {
                        //Check if relationship exists.
                        var exisistingRel = targetWorksheet._drawings.Part.GetRelationshipsByType(srcsRel.RelationshipType).Where(x => x.TargetUri == srcsRel.TargetUri).FirstOrDefault();
                        //Create new relation id if no relation exists or if it's a different worksheet. Otherwise asign the existing relationship Id
                        if (exisistingRel == null || targetWorksheet != _drawings.Worksheet)
                        {
                            newRel = targetWorksheet._drawings.Part.CreateRelationshipFromCopy(srcsRel);
                            relNode.Value = newRel.Id;
                        }
                        else
                        {
                            relNode.Value = exisistingRel.Id;
                        }
                    }
                }
            }
            if (!isGroupShape)
            {
                //Set New id on copied picture.
                var pic = GetDrawing(targetWorksheet._drawings, drawNode) as ExcelPicture;
                pic.SetNewId(++targetWorksheet.Drawings._nextDrawingId);
                pic.Name = targetWorksheet._drawings.GetUniqueDrawingName(this.Name);
            }
            return drawNode;
        }

        private XmlNode CopyShape(ExcelChartStandard targetChart, bool isGroupShape = false, XmlNode groupDrawNode = null)
        {
            XmlNode drawNode = null;
            if (isGroupShape && groupDrawNode != null)
            {
                drawNode = groupDrawNode;
                groupDrawNode.SelectSingleNode("cdr:nvSpPr/cdr:cNvPr", targetChart.Drawings.NameSpaceManager).Attributes["id"].Value = (++targetChart.Drawings._nextDrawingId).ToString();
                groupDrawNode.SelectSingleNode("cdr:nvSpPr/cdr:cNvPr", targetChart.Drawings.NameSpaceManager).Attributes["name"].Value = targetChart.Drawings.GetUniqueDrawingName(this.Name);
            }
            else
            {
                drawNode = targetChart.Drawings.CreateDrawingXmlChartDrawings(targetChart);
                drawNode.InnerXml = TopNode.InnerXml;
                var targetShape = GetDrawing(targetChart.Drawings._drawings, drawNode, DrawingsCollectionType.Chart) as ExcelShape;
                targetShape.Id = ++targetChart.Drawings._nextDrawingId;
                targetShape.Name = targetChart.Drawings.GetUniqueDrawingName(this.Name);
            }
            return drawNode;
        }

        private XmlNode CopyShape(ExcelWorksheet worksheet, bool isGroupShape = false, XmlNode groupDrawNode = null)
        {
            var sourceShape = this as ExcelShape;
            XmlNode drawNode = null;
            if (isGroupShape && groupDrawNode != null)
            {
                drawNode = groupDrawNode;
                groupDrawNode.SelectSingleNode("xdr:nvSpPr/xdr:cNvPr", worksheet._drawings.NameSpaceManager).Attributes["id"].Value = (++worksheet.Drawings._nextDrawingId).ToString();
                groupDrawNode.SelectSingleNode("xdr:nvSpPr/xdr:cNvPr", worksheet._drawings.NameSpaceManager).Attributes["name"].Value = worksheet._drawings.GetUniqueDrawingName(sourceShape.Name);
            }
            else
            {
                //Create node in drawing.xml
                drawNode = worksheet.Drawings.CreateDocumentAndTopNode(CellAnchor, false);
                drawNode.InnerXml = TopNode.InnerXml;
                //Asign new id
                var targetShape = GetDrawing(worksheet._drawings, drawNode) as ExcelShape;
                targetShape.Id = ++worksheet.Drawings._nextDrawingId;
                targetShape.Name = worksheet._drawings.GetUniqueDrawingName(sourceShape.Name);
            }
            //Copy Blip Fill
            WorksheetCopyHelper.CopyBlipFillDrawing(worksheet, worksheet._drawings.Part, worksheet._drawings.DrawingXml, this, sourceShape.Fill, worksheet._drawings.Part.Uri);
            return drawNode;
        }

        internal ExcelAddressBase GetAddress()
        {
            GetFromBounds(out int fromRow, out _, out int fromCol, out _);
            GetToBounds(out int toRow, out _, out int toCol, out _);
            return new ExcelAddress(fromRow + 1, fromCol + 1, toRow + 1, toCol + 1);
        }
    }
}
