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
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Packaging.Ionic;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using OfficeOpenXml.Utils.TypeConversion;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Base class for drawings. 
    /// Drawings are Charts, Shapes and Pictures.
    /// </summary>
    public class ExcelDrawing : XmlHelper, IDisposable, IPictureContainer
    {
        internal ExcelDrawings _drawings;
        internal ExcelGroupShape _parent;
        internal string _topPath, _nvPrPath, _hyperLinkPath;
        internal string _topPathUngrouped, _nvPrPathUngrouped;
        internal int _id;
        internal const float STANDARD_DPI = 96;
        /// <summary>
        /// Ratio between EMU and Pixels
        /// </summary>
        public const int EMU_PER_PIXEL = 9525;
        /// <summary>
        /// Ratio between EMU and Points
        /// </summary>
        public const int EMU_PER_POINT = 12700;
        public const int EMU_PER_CM = 360000;
        public const int EMU_PER_MM = 3600000;
        public const int EMU_PER_US_INCH = 914400;
        public const int EMU_PER_PICA = EMU_PER_US_INCH / 6;

        internal double _width = double.MinValue, _height = double.MinValue, _top = double.MinValue, _left = double.MinValue;
        internal static readonly string[] _schemaNodeOrderSpPr = new string[] { "xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "pattFill", "grpFill", "blipFill", "ln", "effectLst", "effectDag", "scene3d", "sp3d" };

        internal protected bool _doNotAdjust = false;
        internal ExcelDrawing(ExcelDrawings drawings, XmlNode node, string topPath, string nvPrPath, ExcelGroupShape parent = null) :
            base(drawings.NameSpaceManager, node)
        {
            _drawings = drawings;
            _parent = parent;
            if (node != null)   //No drawing, chart xml only. This currently happends when created from a chart template
            {
                TopNode = node;
                
                if(DrawingType==eDrawingType.Control || drawings.Worksheet.Workbook._nextDrawingId >= 1025)
                {
                    _id = drawings.Worksheet._nextControlId++;
                }
                else
                {
                    _id = drawings.Worksheet.Workbook._nextDrawingId++;
                }

                AddSchemaNodeOrder(new string[] { "from", "pos", "to", "ext", "pic", "graphicFrame", "sp", "cxnSp ","grpSp", "nvSpPr", "nvCxnSpPr", "nvGraphicFramePr", "spPr", "style", "AlternateContent", "clientData" }, _schemaNodeOrderSpPr);
                _topPathUngrouped = topPath;
                _nvPrPathUngrouped = nvPrPath;
                if (_parent == null)
                {
                    AdjustXPathsForGrouping(false);
                    CellAnchor = GetAnchorFromName(node.LocalName);
                    SetPositionProperties(drawings, node);
                    GetPositionSize();                                  //Get the drawing position and size, so we can adjust it upon save, if the normal font is changed 

                    string relID = GetXmlNodeString(_hyperLinkPath + "/@r:id");
                    if (!string.IsNullOrEmpty(relID))
                    {
                        HypRel = drawings.Part.GetRelationship(relID);
                        if (HypRel.TargetUri == null)
                        {
                            if (!string.IsNullOrEmpty(HypRel.Target))
                            {
                                Hyperlink = new ExcelHyperLink(HypRel.Target.Substring(1), "");
                            }
                        }
                        else
                        {
                            if (HypRel.TargetUri.IsAbsoluteUri)
                            {
                                Hyperlink = new ExcelHyperLink(HypRel.TargetUri.AbsoluteUri);
                            }
                            else
                            {
                                Hyperlink = new ExcelHyperLink(HypRel.TargetUri.OriginalString, UriKind.Relative);
                            }
                        }
                        ((ExcelHyperLink)Hyperlink).ToolTip = GetXmlNodeString(_hyperLinkPath + "/@tooltip");
                    }
                }
                else
                {
                    AdjustXPathsForGrouping(true);
                    SetPositionProperties(drawings, node);
                    GetPositionSize();                                  //Get the drawing position and size, so we can adjust it upon save, if the normal font is changed 
                }
            }   
        }

        internal virtual void AdjustXPathsForGrouping(bool group)
        {
            if(group)
            {
                _topPath = _topPathUngrouped.IndexOf('/') > 0 ? _topPathUngrouped.Substring(_topPathUngrouped.IndexOf('/')+1) : "";
                if(_topPath=="")
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

        private void SetPositionProperties(ExcelDrawings drawings, XmlNode node)
        {
            if (_parent == null) //Top level drawing
            {
                XmlNode posNode = node.SelectSingleNode("xdr:from", drawings.NameSpaceManager);
                if (posNode != null)
                {
                    From = new ExcelPosition(drawings.NameSpaceManager, posNode, GetPositionSize);
                }
                else
                {
                    posNode = node.SelectSingleNode("xdr:pos", drawings.NameSpaceManager);
                    if (posNode != null)
                    {
                        Position = new ExcelDrawingCoordinate(drawings.NameSpaceManager, posNode, GetPositionSize);
                    }
                }
                posNode = node.SelectSingleNode("xdr:to", drawings.NameSpaceManager);
                if (posNode != null)
                {
                    To = new ExcelPosition(drawings.NameSpaceManager, posNode, GetPositionSize);
                }
                else
                {
                    To = null;
                    posNode = node.SelectSingleNode("xdr:ext", drawings.NameSpaceManager);
                    if (posNode != null)
                    {
                        Size = new ExcelDrawingSize(drawings.NameSpaceManager, posNode, GetPositionSize);
                    }
                }
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

        private XmlNode GetXFrameNode(XmlNode node, string child)
        {
            if(node.LocalName == "AlternateContent")
            {
                node = node.ChildNodes[0].ChildNodes[0];
            }
            if (node.LocalName == "grpSp")
            {
                return node.SelectSingleNode($"xdr:grpSpPr/a:xfrm/{child}", NameSpaceManager);
            }
            else if (node.LocalName == "graphicFrame")
            {
                return node.SelectSingleNode($"xdr:xfrm/{child}", NameSpaceManager);
            }
            else
            {
                return node.SelectSingleNode($"xdr:spPr/a:xfrm/{child}", NameSpaceManager);
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
                case "absoluteAnchor":
                    return eEditAs.Absolute;
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
                    return GetXmlNodeString(_nvPrPath+"/@name");
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
                    if (_parent!=null && DrawingType == eDrawingType.Control)
                    {
                        return ((ExcelControl)this).GetCellAnchorFromWorksheetXml();
                    }
                    if (CellAnchor == eEditAs.TwoCell)
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
                if(_parent!=null)
                {
                    if(DrawingType==eDrawingType.Control)
                    {
                        ((ExcelControl)this).SetCellAnchor(value);
                    }
                    else
                    {
                        throw (new InvalidOperationException("EditAs can't be set when a drawing is a part of a group."));
                    }
                }
                else if (CellAnchor == eEditAs.TwoCell)
                {
                    string s = value.ToString();
                    SetXmlNodeString("@editAs", s.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + s.Substring(1, s.Length - 1));
                }
                else if(CellAnchor!=value)
                {
                    throw (new InvalidOperationException("EditAs can only be set when CellAnchor is set to TwoCellAnchor"));
                }
            }
        }
        const string lockedPath="xdr:clientData/@fLocksWithSheet";
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
        /// Otherwise this propery is set to null
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
            private set;
        }
        /// <summary>
        /// The extent of the shape, if the shape is of the one- or absolute- anchor type.
        /// Otherwise this propery is set to null
        /// </summary>
        public ExcelDrawingSize Size
        {
            get;
            private set;
        }
        /// <summary>
        /// Bottom right position
        /// </summary>
        public ExcelPosition To { get; private set; } = null;
        Uri _hyperLink=null;
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
                    if(value is ExcelHyperLink el && !string.IsNullOrEmpty(el.ReferenceAddress))
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
        /// <returns>The Drawing object</returns>
        internal static ExcelDrawing GetDrawing(ExcelDrawings drawings, XmlNode node)
        {
            if (node.ChildNodes.Count < 3) return null; //Invalid formatted anchor node, ignore
            XmlElement drawNode = (XmlElement)node.ChildNodes[2];
            return GetDrawingFromNode(drawings, node, drawNode);
        }

        internal static ExcelDrawing GetDrawingFromNode(ExcelDrawings drawings, XmlNode node, XmlElement drawNode, ExcelGroupShape parent=null)
        {
            switch (drawNode.LocalName)
            {
                case "sp":
                    return GetShapeOrControl(drawings, node, drawNode, parent);
                case "pic":
                    return new ExcelPicture(drawings, node, parent);
                case "graphicFrame":
                    return ExcelChart.GetChart(drawings, node, parent);
                case "grpSp":
                    return new ExcelGroupShape(drawings, node, parent);
                case "cxnSp":
                    return new ExcelConnectionShape(drawings, node, parent);
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
                                else if (choice.ChildNodes.Count > 0 && choice.FirstChild.LocalName=="sp")
                                {
                                    return GetShapeOrControl(drawings, node, (XmlElement)choice.FirstChild, parent);
                                }
                                break;

                        }
                    }
                    break;
            }
            return new ExcelDrawing(drawings, node, "", "");
       }

        private static ExcelDrawing GetShapeOrControl(ExcelDrawings drawings, XmlNode node, XmlElement drawNode, ExcelGroupShape parent)
        {
            var shapeId = GetControlShapeId(drawNode, drawings.NameSpaceManager);
            var control = drawings.Worksheet.Controls.GetControlByShapeId(shapeId);
            if (control != null)
            {
                return ControlFactory.GetControl(drawings, drawNode, control, parent);
            }
            else
            {
                return new ExcelShape(drawings, node, parent);
            }
        }

        private static int GetControlShapeId(XmlElement drawNode, XmlNamespaceManager nameSpaceManager)
        {
            var idNode = drawNode.SelectSingleNode("xdr:nvSpPr/xdr:cNvPr/@id", nameSpaceManager);
            if(idNode!=null)
            {
                return int.Parse(idNode.Value);
            }
            return -1;
        }

        internal int Id
        {
            get { return _id; }
        }
        #region "Internal sizing functions"
        internal int GetPixelLeft()
        {
            int pix;
            if (CellAnchor == eEditAs.Absolute)
            {
                pix = Position.X / EMU_PER_PIXEL;
            }
            else
            {
                ExcelWorksheet ws = _drawings.Worksheet;
                decimal mdw = ws.Workbook.MaxFontWidth;

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
            int pix;
            if (CellAnchor == eEditAs.Absolute)
            {
                pix = Position.Y / EMU_PER_PIXEL;
            }
            else
            {
                pix = 0;
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
            double pix;
            if (CellAnchor == eEditAs.TwoCell)
            {
                ExcelWorksheet ws = _drawings.Worksheet;
                decimal mdw = ws.Workbook.MaxFontWidth;

                pix = -From.ColumnOff / (double)EMU_PER_PIXEL;
                for (int col = From.Column + 1; col <= To.Column; col++)
                {
                    pix += (double)decimal.Truncate(((256 * ws.GetColumnWidth(col) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
                }
                pix += Convert.ToDouble(To.ColumnOff) / (double)EMU_PER_PIXEL;
            }
            else
            {
                pix = Size.Width / (double)EMU_PER_PIXEL;
            }
            return pix;
        }
        internal double GetPixelHeight()
        {
            double pix;
            if (CellAnchor == eEditAs.TwoCell)
            {
                ExcelWorksheet ws = _drawings.Worksheet;

                pix = -(From.RowOff / (double)EMU_PER_PIXEL);
                for (int row = From.Row + 1; row <= To.Row; row++)
                {
                    pix += ws.GetRowHeight(row) / 0.75;
                }
                pix += Convert.ToDouble(To.RowOff) / EMU_PER_PIXEL;
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
                Position.Y = (int)(pixels * EMU_PER_PIXEL);
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
            decimal mdw = ws.Workbook.MaxFontWidth;
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
                Position.X = (int)(pixels * EMU_PER_PIXEL);
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
            decimal mdw = ws.Workbook.MaxFontWidth;
            double prevPix = 0;
            double pix = (int)decimal.Truncate(((256 * ws.GetColumnWidth(1) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
            int col = 2;

            while (pix < pixels)
            {
                prevPix = pix;
                pix += (int)decimal.Truncate(((256 * ws.GetColumnWidth(col++) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
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
            if (CellAnchor == eEditAs.TwoCell)
            {
                _doNotAdjust = true;
                GetToRowFromPixels(pixels,  out int toRow, out int pixOff);
                To.Row = toRow;
                To.RowOff = pixOff;
                _doNotAdjust = false;
            }
            else
            {
                Size.Height = (long)Math.Round(pixels * EMU_PER_PIXEL);
            }
        }

        internal void GetToRowFromPixels(double pixels, out int toRow, out int rowOff, int fromRow=-1, int fromRowOff=-1)
        {
            if(fromRow<0)
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
            if (CellAnchor == eEditAs.TwoCell)
            {
                _doNotAdjust = true;
                GetToColumnFromPixels(pixels, out int col, out double pixOff);

                To.Column = col - 2;
                To.ColumnOff = (int)(pixOff * EMU_PER_PIXEL);
                _doNotAdjust = false;
            }
            else
            {
                Size.Width = (int)Math.Round(pixels * EMU_PER_PIXEL);
            }
        }

        internal void GetToColumnFromPixels(double pixels, out int col, out double prevRowOff, int fromColumn = -1, int fromColumnOff = -1)
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            decimal mdw = ws.Workbook.MaxFontWidth;
            if(fromColumn<0)
            {
                fromColumn = From.Column;
                fromColumnOff = From.ColumnOff;
            }
            double pixOff = pixels - (double)(decimal.Truncate(((256 * ws.GetColumnWidth(fromColumn + 1) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw) - fromColumnOff / EMU_PER_PIXEL);
            prevRowOff = fromColumnOff / EMU_PER_PIXEL + pixels;
            col = fromColumn + 2;
            while (pixOff >= 0)
            {
                prevRowOff = pixOff;
                pixOff -= (double)decimal.Truncate(((256 * ws.GetColumnWidth(col++) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
            }
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
            SetPosition(PixelTop, PixelLeft, true);
        }
        internal void SetPosition(int PixelTop, int PixelLeft, bool adjustChildren)
        {
            _doNotAdjust = true;
            if (_width == int.MinValue)
            {
                _width = GetPixelWidth();
                _height = GetPixelHeight();
            }
            if(adjustChildren && DrawingType == eDrawingType.GroupShape)
            {
                if(_left== int.MinValue)
                {
                    _left = GetPixelLeft();
                    _top = GetPixelTop();
                }
                var grp = (ExcelGroupShape)this;
                foreach(var d in grp.Drawings)
                {
                    d.SetPosition((int)(d._top + (PixelTop - _top)), (int)(d._left + (PixelLeft - _left)));
                }
            }
            SetPixelTop(PixelTop);
            SetPixelLeft(PixelLeft);

            SetPixelWidth(_width);
            SetPixelHeight(_height);
            _doNotAdjust = false;
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
        /// This will change the cell anchor type without modifiying the position and size.
        /// </summary>
        /// <param name="type">The cell anchor type to change to</param>
        public void ChangeCellAnchor(eEditAs type)
        {
            if(DrawingType==eDrawingType.Control)
            {
                throw new InvalidOperationException("Controls can't change CellAnchor. Must be TwoCell anchor. Please use EditAs property instead.");
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

        private void ChangeCellAnchorTypeInternal(eEditAs type)
        {
            if (type != CellAnchor)
            {
                CellAnchor = type;
                RenameNode(TopNode, "xdr", $"{type.ToEnumString()}Anchor");
                CleanupPositionXml();
                SetPositionProperties(_drawings, TopNode);
                CellAnchorChanged();
            }
        }

        internal virtual void CellAnchorChanged()
        {
            
        }

        private void CleanupPositionXml()
        {
            switch(CellAnchor)
            {
                case eEditAs.OneCell:
                    DeleteNode("xdr:to");
                    DeleteNode("xdr:pos");
                    CreateNode("xdr:from");
                    CreateNode("xdr:ext");
                    break;
                case eEditAs.Absolute:
                    DeleteNode("xdr:to");
                    DeleteNode("xdr:from"); 
                    CreateNode("xdr:pos");
                    CreateNode("xdr:ext");
                    break;
                default:
                    DeleteNode("xdr:pos");
                    DeleteNode("xdr:ext");
                    CreateNode("xdr:from");
                    CreateNode("xdr:to");
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
        }
        /// <summary>
        /// Set size in Percent.
        /// Note that resizing columns / rows after using this function will effect the size of the drawing
        /// </summary>
        /// <param name="Percent"></param>
        public virtual void SetSize(int Percent)
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
        /// <summary>
        /// Set size in pixels
        /// Note that resizing columns / rows after using this function will effect the size of the drawing
        /// </summary>
        /// <param name="PixelWidth">Width in pixels</param>
        /// <param name="PixelHeight">Height in pixels</param>
        public void SetSize(int PixelWidth, int PixelHeight)
        {
            _doNotAdjust = true;
            _width = PixelWidth;
            _height = PixelHeight;
            SetPixelWidth(PixelWidth);
            SetPixelHeight(PixelHeight);
            _doNotAdjust = false;
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
            foreach(var d in drawing)
            {
                ExcelGroupShape.Validate(d, _drawings, grp);
                if (d._parent != null) grp = d._parent;
            }
            if (grp == null)
            {
                grp = _drawings.AddGroupDrawing();
            }
            
            grp.Drawings.AddDrawing(this);

            foreach (var d in drawing)
            {
                grp.Drawings.AddDrawing(d);
            }

            grp.SetPositionAndSizeFromChildren();
            return grp;
        }
        internal XmlElement GetFrmxNode(XmlNode node)
        {
            if(node.LocalName == "AlternateContent")
            {
                node = node.FirstChild.FirstChild;
            }

            if(node.LocalName == "sp" || node.LocalName == "pic" || node.LocalName == "cxnSp")
            {
                return (XmlElement)CreateNode(node, "xdr:spPr/a:xfrm");
            }
            else if(node.LocalName == "graphicFrame")
            {
                return (XmlElement)CreateNode(node, "xdr:xfrm"); 
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
        public void UnGroup(bool ungroupThisItemOnly=true)
        {
            if(_parent==null)
            {
                throw new InvalidOperationException("Cannot ungroup this drawing. This drawing is not part of a group");
            }
            if(ungroupThisItemOnly)
            {
                _parent.Drawings.Remove(this);
            }
            else
            {
                _parent.Drawings.Clear();
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
            if (EditAs==eEditAs.Absolute)
            {
                SetPixelLeft(_left);
                SetPixelTop(_top);
            }
            if(EditAs == eEditAs.Absolute || EditAs == eEditAs.OneCell)
            {
                SetPixelHeight(_height);
                SetPixelWidth(_width);
            }
            _drawings.Worksheet.Workbook._package.DoAdjustDrawings = true;
        }
        internal protected XmlElement CreateShapeNode()
        {
            XmlElement shapeNode = TopNode.OwnerDocument.CreateElement("xdr", "sp", ExcelPackage.schemaSheetDrawings);
            shapeNode.SetAttribute("macro", "");
            shapeNode.SetAttribute("textlink", "");
            TopNode.AppendChild(shapeNode);
            return shapeNode;
        }
        internal protected XmlElement CreateClientData()
        {
            XmlElement clientDataNode = TopNode.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings);
            clientDataNode.SetAttribute("fPrintsWithSheet", "0");
            TopNode.ChildNodes[2].ChildNodes[0].ChildNodes[0].AppendChild(clientDataNode);
            TopNode.AppendChild(clientDataNode);
            return clientDataNode;
        }
        string IPictureContainer.ImageHash { get; set; }
        Uri IPictureContainer.UriPic { get; set; }
        Packaging.ZipPackageRelationship IPictureContainer.RelPic { get; set; }
        IPictureRelationDocument IPictureContainer.RelationDocument => _drawings as IPictureRelationDocument;
    }
}
