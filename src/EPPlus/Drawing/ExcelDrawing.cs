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
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.Worksheet;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
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
        /// The ratio between EMU and milimeters
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

        private XmlNode GetXFrameNode(XmlNode node, string child)
        {
            if(node.LocalName == "AlternateContent")
            {
                node = node.GetChildAtPosition(0).GetChildAtPosition(0);
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
            XmlElement drawNode = (XmlElement)node.GetChildAtPosition(2);
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
                                else if (choice.ChildNodes.Count > 0)
                                {
                                    if (choice.FirstChild.LocalName == "sp")
                                    {
                                        return GetShapeOrControl(drawings, node, (XmlElement)choice.FirstChild, parent);
                                    }
                                    else if(choice.FirstChild.LocalName == "grpSp")
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

                var w = (double)decimal.Truncate(((256 * ws.GetColumnWidth(To.Column + 1) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
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
            double pix;
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

        internal void GetToColumnFromPixels(double pixels, out int col, out int colOff, int fromColumn = -1, int fromColumnOff = -1)
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            decimal mdw = ws.Workbook.MaxFontWidth;
            if(fromColumn<0)
            {
                fromColumn = From.Column;
                fromColumnOff = From.ColumnOff;
            }
            double pixOff = pixels - (double)(decimal.Truncate(((256 * ws.GetColumnWidth(fromColumn + 1) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw) - fromColumnOff / EMU_PER_PIXEL);
            double offset = (double)fromColumnOff / EMU_PER_PIXEL + pixels;
            col = fromColumn + 2;
            while (pixOff >= 0)
            {
                offset = pixOff;
                pixOff -= (double)decimal.Truncate(((256 * ws.GetColumnWidth(col++) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
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
        internal void SetCellAnchorFromNode()
        {
            if(TopNode.LocalName== "twoCellAnchor")
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
            UpdatePositionAndSizeXml();
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
            UpdatePositionAndSizeXml();
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
        internal void UpdatePositionAndSizeXml()
        {
            From?.UpdateXml();
            To?.UpdateXml();
            Size?.UpdateXml();
            Position?.UpdateXml();
        }


        internal XmlElement CreateShapeNode()
        {
            XmlElement shapeNode = TopNode.OwnerDocument.CreateElement("xdr", "sp", ExcelPackage.schemaSheetDrawings);
            shapeNode.SetAttribute("macro", "");
            shapeNode.SetAttribute("textlink", "");
            TopNode.AppendChild(shapeNode);
            return shapeNode;
        }
        internal XmlElement CreateClientData()
        {
            XmlElement clientDataNode = TopNode.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings);
            clientDataNode.SetAttribute("fPrintsWithSheet", "0");
            TopNode.GetChildAtPosition(2).GetChildAtPosition(0).GetChildAtPosition(0).AppendChild(clientDataNode);
            TopNode.AppendChild(clientDataNode);
            return clientDataNode;
        }


        public void Copy(ExcelWorksheet worksheet, int row, int col, int rowOffset = int.MinValue, int colOffset = int.MinValue)
        {
            XmlNode drawNode = null;
            if(rowOffset == int.MinValue)
            {
                rowOffset = From.RowOff / 9525;
            }
            if(colOffset == int.MinValue)
            {
                colOffset = From.ColumnOff / 9525;
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
        }

        private XmlNode CopyGroupShape(ExcelWorksheet worksheet)
        {
            //Create node in drawing.xml
            var drawNode = worksheet.Drawings.CreateDocumentAndTopNode(CellAnchor, false);
            drawNode.InnerXml = TopNode.InnerXml;
            CopyGroupShape(worksheet, this, drawNode.ChildNodes[2]);
            return drawNode;
        }

        private void CopyGroupShape(ExcelWorksheet targetWorksheet, ExcelDrawing sourceDrawing, XmlNode targetDrawNode)
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
            else if(sourceDrawing is ExcelTableSlicer tSlicer)
            {
                sourceDrawing.CopySlicer(targetWorksheet, true, targetDrawNode);
            }
            else if (sourceDrawing is ExcelPivotTableSlicer ptSlicer)
            {
                sourceDrawing.CopySlicer(targetWorksheet, true, targetDrawNode);
            }
            else if (sourceDrawing is ExcelGroupShape groupShape)
            {
                int nodeIndex = 2;
                for (int j = 0; j < groupShape.Drawings.Count; j++)
                {
                    //börja på 2 men child nodes måste inkrementeras med 1 varje varv så vi kikar på nästa nod!
                    CopyGroupShape(targetWorksheet, groupShape.Drawings[j], targetDrawNode.ChildNodes[nodeIndex++]);
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
            if(isGroupShape)
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
            if(drawNodeName == null && isGroupShape)
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
            if (wsSlicerNode == null) {
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
            else if(this is ExcelPivotTableSlicer epts)
            {
                xmlSource = _drawings.Worksheet.SlicerXmlSources._list.Find(x => x == epts._xmlSource);
                name = epts.Name;
            }
            //If different drawings create a new xml. (Maybe check for exsisting xml in new drawings and append instead)
            if (_drawings != worksheet._drawings)
            {
                if (isNewPart)
                {
                    XmlHelper.LoadXmlSafe(xmlTarget, "<slicers xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" mc:Ignorable=\"x xr10\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" />", Encoding.UTF8);
                }
                else
                {
                    xmlTarget = worksheet.SlicerXmlSources._list.Find(x => x.Type == eSlicerSourceType.Table).XmlDocument; //håller en kopi, need to skriv ref...
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

        private XmlNode CopyChart(ExcelWorksheet worksheet, bool isGroupShape = false, XmlNode groupDrawNode = null)
        {
            XmlNode drawNode = null;
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
            var relNode =  drawNode.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart/@r:id", NameSpaceManager);
            if(relNode == null)
            {
                relNode = drawNode.SelectSingleNode("a:graphic/a:graphicData/c:chart/@r:id", NameSpaceManager);
            }
            if (relNode != null && _drawings.Part.RelationshipExists(relNode.Value))
            {
                var origialChart = this as ExcelChart;
                WorksheetCopyHelper.CopyChartRelations(origialChart, worksheet, worksheet._drawings.Part, worksheet._drawings.DrawingXml, _drawings.Worksheet);
                //Update the copied charts id and name
                if(isGroupShape)
                {
                    var chartAttr = groupDrawNode.SelectSingleNode("xdr:nvGraphicFramePr/xdr:cNvPr", worksheet._drawings.NameSpaceManager);
                    chartAttr.Attributes["name"].Value = worksheet._drawings.GetUniqueDrawingName(origialChart.Name);
                    chartAttr.Attributes["id"].Value = (++origialChart._id).ToString();
                }
                else
                {
                    var chartcopy = ExcelChart.GetChart(worksheet._drawings, drawNode);
                    chartcopy.Name = worksheet._drawings.GetUniqueDrawingName(origialChart.Name);
                    chartcopy._id = ++origialChart._id;
                }

            }
            return drawNode;
        }

        private XmlNode CopyPicture(ExcelWorksheet worksheet, bool isGroupShape = false, XmlNode groupDrawNode = null)
        {
            XmlNode drawNode = null;
            if (isGroupShape && groupDrawNode != null)
            {
                drawNode = groupDrawNode;
                groupDrawNode.SelectSingleNode("xdr:nvPicPr/xdr:cNvPr", worksheet._drawings.NameSpaceManager).Attributes["id"].Value = (++worksheet.Workbook._nextDrawingId).ToString();
            }
            else
            {
                //Create node in drawing.xml
                drawNode = worksheet.Drawings.CreateDocumentAndTopNode(CellAnchor, false);
                drawNode.InnerXml = TopNode.InnerXml;
            }
            //If same drawings object, we are done.
            if (worksheet._drawings != _drawings)
            {
                //Get the relation node
                var relNode = drawNode.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager);
                if(relNode == null)
                {
                    relNode = drawNode.SelectSingleNode("xdr:blipFill/a:blip/@r:embed", NameSpaceManager);
                }
                if (relNode != null && _drawings.Part.RelationshipExists(relNode.Value))
                {
                    var rel = _drawings.Part.GetRelationship(relNode.Value);
                    //Copy image file to new workbook if target worksheet is in a different workbook.
                    if (worksheet.Workbook != _drawings.Worksheet.Workbook)
                    {
                        var uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                        var imagePart = _drawings.Worksheet.Workbook._package.ZipPackage.GetPart(uri);
                        var imageStream = (MemoryStream)imagePart.GetStream(FileMode.Open, FileAccess.Read);
                        var image = new byte[imageStream.Length];
                        imageStream.Seek(0, SeekOrigin.Begin);
                        imageStream.Read(image, 0, (int)imageStream.Length);
                        var imageInfo = worksheet.Workbook._package.PictureStore.GetImageInfo(image);
                        if (imageInfo == null)
                        {
                            var copyPart = worksheet.Workbook._package.ZipPackage.CreatePart(uri, imagePart.ContentType);
                            var copyStream = (MemoryStream)copyPart.GetStream(FileMode.Create, FileAccess.Write);
                            copyStream.Write(image, 0, image.Length);
                        }
                        else
                        {
                            rel.TargetUri = imageInfo.Uri;
                        }
                    }
                    //Check if relationship exists.
                    var exisistingRel = worksheet._drawings.Part.GetRelationshipsByType(rel.RelationshipType).Where(x => x.Target == rel.Target).FirstOrDefault();
                    //Create new relation id if no relation exsist or if it's a different worksheet. Otherwise asign the exsisting relationship Id
                    if (exisistingRel == null || worksheet != _drawings.Worksheet)
                    {
                        var newRel = worksheet._drawings.Part.CreateRelationshipFromCopy(rel);
                        relNode.Value = newRel.Id;
                    }
                    else
                    {
                        relNode.Value = exisistingRel.Id;
                    }
                }
            }
            if (!isGroupShape)
            {
                //Set New id on copied picture.
                var pic = GetDrawing(worksheet._drawings, drawNode) as ExcelPicture;
                pic.SetNewId(++worksheet.Workbook._nextDrawingId);
                pic.Name = worksheet._drawings.GetUniqueDrawingName(this.Name);
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
                groupDrawNode.SelectSingleNode("xdr:nvSpPr/xdr:cNvPr", worksheet._drawings.NameSpaceManager).Attributes["id"].Value = (++worksheet.Workbook._nextDrawingId).ToString();
                groupDrawNode.SelectSingleNode("xdr:nvSpPr/xdr:cNvPr", worksheet._drawings.NameSpaceManager).Attributes["name"].Value = worksheet._drawings.GetUniqueDrawingName(sourceShape.Name);
            }
            else
            {
                //Create node in drawing.xml
                drawNode = worksheet.Drawings.CreateDocumentAndTopNode(CellAnchor, false);
                drawNode.InnerXml = TopNode.InnerXml;
                //Asign new id
                var targetShape = GetDrawing(worksheet._drawings, drawNode) as ExcelShape;
                targetShape._id = ++worksheet.Workbook._nextDrawingId;
                targetShape.Name = worksheet._drawings.GetUniqueDrawingName(sourceShape.Name);
            }
            //Copy Blip Fill
            WorksheetCopyHelper.CopyBlipFillDrawing(worksheet, worksheet._drawings.Part, worksheet._drawings.DrawingXml, this, sourceShape.Fill, worksheet._drawings.Part.Uri);
            return drawNode;
        }
    }
}
