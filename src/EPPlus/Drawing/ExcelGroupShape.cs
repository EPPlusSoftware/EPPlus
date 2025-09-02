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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Utils.XML;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Xml;
using System.Xml.Linq;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// A collection of child drawings to a group drawing
    /// </summary>
    public class ExcelDrawingsGroup : IEnumerable<ExcelDrawing>, IDisposable
    {
        private ExcelGroupShape _parent;
        internal Dictionary<string, int> _drawingNames;
        private List<ExcelDrawing> _groupDrawings;
        internal DrawingsCollectionType _drawingsCollectionType;
        XmlNamespaceManager _nsm;
        XmlNode _topNode;
        internal ExcelDrawingsGroup(ExcelGroupShape parent, XmlNamespaceManager nsm, XmlNode topNode, DrawingsCollectionType drawingsCollectionType = DrawingsCollectionType.Worksheet)
        {
            _parent = parent;
            _nsm = nsm;
            _topNode = topNode;
            _drawingNames = new Dictionary<string, int>();
            _drawingsCollectionType = drawingsCollectionType;
            AddDrawings();
        }
        private void AddDrawings()
        {
            _groupDrawings = new List<ExcelDrawing>();
            foreach (XmlNode node in _topNode.ChildNodes)
            {

                if (node.LocalName != "nvGrpSpPr" && node.LocalName != "grpSpPr")
                {
                    var grpDraw = ExcelDrawing.GetDrawingFromNode(_parent._drawings, node, (XmlElement)node, _parent, _drawingsCollectionType);
                    _groupDrawings.Add(grpDraw);
                    if (_drawingNames.ContainsKey(grpDraw.Name) == false)
                    {
                        _drawingNames.Add(grpDraw.Name, _groupDrawings.Count - 1);
                    }
                }
            }
        }
        /// <summary>
        /// Adds a drawing to the group
        /// </summary>
        /// <param name="drawing"></param>
        public void Add(ExcelDrawing drawing)
        {
            CheckNotDisposed();
            AddDrawing(drawing);
            drawing.ParentGroup.SetPositionAndSizeFromChildren();
        }

        private void CheckNotDisposed()
        {
            if (_topNode == null)
            {
                throw (new ObjectDisposedException("This group drawing has been disposed."));
            }
        }

        internal void AddDrawing(ExcelDrawing drawing)
        {
            if (drawing._parent == _parent) return; //This drawing is already added to the group, exit

            ExcelGroupShape.Validate(drawing, drawing._drawings, _parent);
            AdjustXmlAndMoveToGroup(drawing);
            ExcelGroupShape.Validate(drawing, _parent._drawings, _parent);
            AppendDrawingNode(drawing.TopNode);
            drawing._parent = _parent;

            _groupDrawings.Add(drawing);
            _drawingNames.Add(drawing.Name, _groupDrawings.Count - 1);
        }

        private void AdjustXmlAndMoveToGroup(ExcelDrawing d)
        {
            d._drawings.RemoveDrawing(d._drawings._drawingsList.IndexOf(d), false);
            var height = d.GetPixelHeight();
            var width = d.GetPixelWidth();
            var top = d.GetPixelTop();
            var left = d.GetPixelLeft();
            var node = d.TopNode.GetChildAtPosition(2);
            XmlElement xFrmNode = d.GetXfrmNode(node);
            if (xFrmNode.ChildNodes.Count == 0)
            {
                d.CreateNode(xFrmNode, "a:off");
                d.CreateNode(xFrmNode, "a:ext");
            }
            var offNode = (XmlElement)xFrmNode.SelectSingleNode("a:off", _nsm);
            var extNode = (XmlElement)xFrmNode.SelectSingleNode("a:ext", _nsm);
            if (d._drawings._collectionType == DrawingsCollectionType.Worksheet)
            {
                offNode.SetAttribute("y", (top * ExcelDrawing.EMU_PER_PIXEL).ToString());
                offNode.SetAttribute("x", (left * ExcelDrawing.EMU_PER_PIXEL).ToString());
                extNode.SetAttribute("cy", Math.Round(height * ExcelDrawing.EMU_PER_PIXEL, 0).ToString());
                extNode.SetAttribute("cx", Math.Round(width * ExcelDrawing.EMU_PER_PIXEL, 0).ToString());
                d.SetGroupChild(offNode, extNode);
            }
            else if (d._drawings._collectionType == DrawingsCollectionType.Chart)
            {
                if (d is not ExcelGroupShape)
                {
                    d.RemoveFromToNodes();
                    d.Position = new ExcelDrawingCoordinate(d.NameSpaceManager, d.TopNode, d.GetPositionSize);
                    d.Size = new ExcelDrawingSize(d.NameSpaceManager, d.TopNode, d.GetPositionSize);
                }
            }

            node.ParentNode.RemoveChild(node);
            if (d.TopNode.ParentNode?.ParentNode?.LocalName == "AlternateContent")
            {
                var containerNode = d.TopNode.ParentNode?.ParentNode;
                d.TopNode.ParentNode.RemoveChild(d.TopNode);
                containerNode.ParentNode.RemoveChild(containerNode);
                containerNode.FirstChild.AppendChild(node);
                node = containerNode;
            }
            else
            {
                d.TopNode.ParentNode.RemoveChild(d.TopNode);
            }

            d.AdjustXPathsForGrouping(true);
            d.TopNode = node;
        }
        private void AdjustXmlAndMoveFromGroup(ExcelDrawing d)
        {
            var height = d.GetPixelHeight();
            var width = d.GetPixelWidth();
            var top = d.GetPixelTop();
            var left = d.GetPixelLeft();
            var xmlDoc = _parent.TopNode.OwnerDocument;
            XmlNode drawingNode;
            if (_parent.TopNode.ParentNode?.ParentNode?.LocalName == "AlternateContent") //Create alternat content above ungrouped drawing.
            {
                drawingNode = _parent.TopNode.ParentNode.ParentNode.CloneNode(false);
                var choiceNode = _parent.TopNode.ParentNode.CloneNode(false);
                drawingNode.AppendChild(choiceNode);
                d.TopNode.ParentNode.RemoveChild(d.TopNode);
                choiceNode.AppendChild(d.TopNode);
                drawingNode = CreateAnchorNode(drawingNode);
                var addBeforeNode = _parent.TopNode.ParentNode.ParentNode;
                addBeforeNode.ParentNode.InsertBefore(drawingNode, addBeforeNode);
            }
            else
            {
                d.TopNode.ParentNode.RemoveChild(d.TopNode);
                drawingNode = CreateAnchorNode(d.TopNode);
                _parent.TopNode.ParentNode.InsertBefore(drawingNode, _parent.TopNode);
            }
            d.AdjustXPathsForGrouping(false);
            d.TopNode = drawingNode;
            d.SetCellAnchorFromNode();
            if (d._drawings._collectionType == DrawingsCollectionType.Chart)
            {
                double x1 = 0, y1 = 0, x2 = 0.1, y2 = 0.1;
                MathHelper.AdjustAspectRatio(d._drawings._screenWidth, d._drawings._screenHeight, ref x1, ref y1, ref x2, ref y2);
                double fromX = left / (d._drawings._screenWidth * ExcelDrawing.EMU_PER_PIXEL);
                double fromY = top / (d._drawings._screenHeight * ExcelDrawing.EMU_PER_PIXEL);
                left = (int)(fromX * d._drawings._screenWidth);
                top = (int)(fromY * d._drawings._screenHeight);

                width = (x2 - x1) * d._drawings._screenWidth;
                height = (y2 - y1) * d._drawings._screenHeight;
            }
            d.SetPosition(top, left);
            d.SetSize((int)width, (int)height);
        }

        private XmlNode CreateAnchorNode(XmlNode drawingNode)
        {
            XmlNode topNode;
            var ix = 3;
            if (drawingNode.LocalName == "AlternateContent")
            {
                var xmlDoc = _topNode.OwnerDocument;
                topNode = _topNode.OwnerDocument.CreateElement("xdr", "twoCellAnchor", ExcelPackage.schemaSheetDrawings);
                var from = _topNode.OwnerDocument.CreateElement("xdr", "from", ExcelPackage.schemaSheetDrawings);
                var to = _topNode.OwnerDocument.CreateElement("xdr", "to", ExcelPackage.schemaSheetDrawings);
                topNode.AppendChild(from);
                topNode.AppendChild(to);

                topNode.AppendChild(drawingNode.ChildNodes[0].ChildNodes[0]);
                drawingNode.ChildNodes[0].PrependChild(topNode);
            }
            else
            {
                topNode = _parent.TopNode.CloneNode(false);
                topNode.AppendChild(_parent.TopNode.GetChildAtPosition(0).CloneNode(true));
                topNode.AppendChild(_parent.TopNode.GetChildAtPosition(1).CloneNode(true));
                topNode.AppendChild(drawingNode);
            }

            while (ix < _parent.TopNode.ChildNodes.Count)
            {
                var nodeToAppend = _parent.TopNode.ChildNodes[ix].CloneNode(true);
                topNode.AppendChild(nodeToAppend);
                ix++;
            }
            return topNode;
        }

        private void AppendDrawingNode(XmlNode drawingNode)
        {
            if (drawingNode.ParentNode?.ParentNode?.LocalName == "AlternateContent")
            {
                _topNode.AppendChild(drawingNode.ParentNode.ParentNode);
            }
            else
            {
                _topNode.AppendChild(drawingNode);
            }
        }

        /// <summary>
        /// Disposes the class
        /// </summary>
        public void Dispose()
        {
            _parent = null;
            _topNode = null;
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count { get { return _groupDrawings.Count; } }
        /// <summary>
        /// Returns the drawing at the specified position.  
        /// </summary>
        /// <param name="PositionID">The position of the drawing. 0-base</param>
        /// <returns></returns>
        public ExcelDrawing this[int PositionID]
        {
            get
            {
                return (_groupDrawings[PositionID]);
            }
        }
        /// <summary>
        /// Returns the drawing matching the specified name
        /// </summary>
        /// <param name="Name">The name of the worksheet</param>
        /// <returns></returns>
        public ExcelDrawing this[string Name]
        {
            get
            {
                if (_drawingNames.ContainsKey(Name))
                {
                    return _groupDrawings[_drawingNames[Name]];
                }
                else
                {
                    return null;
                }
            }
        }
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<ExcelDrawing> GetEnumerator()
        {
            return _groupDrawings.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _groupDrawings.GetEnumerator();
        }

        /// <summary>
        /// Removes the <see cref="ExcelDrawing"/> from the group
        /// </summary>
        /// <param name="drawing">The drawing to remove</param>
        public void Remove(ExcelDrawing drawing)
        {
            CheckNotDisposed();
            _groupDrawings.Remove(drawing);
            AdjustXmlAndMoveFromGroup(drawing);
            var ix = _parent._drawings._drawingsList.IndexOf(_parent);
            _parent._drawings._drawingsList.Insert(ix, drawing);

            //Remove 
            if (_parent.Drawings.Count == 0)
            {
                _parent._drawings._drawingsList.Remove(_parent);
                _parent._drawings._drawingNames.Remove(_parent.Name);
            }
            _parent._drawings.ReIndexNames(ix, 1);
            drawing._parent = null;
            if (_parent._collectionType == DrawingsCollectionType.Chart)
            {
                _parent.SetPositionAndSizeFromChildren();
            }
        }
        /// <summary>
        /// Removes all children drawings from the group.
        /// </summary>
        public void Clear()
        {
            CheckNotDisposed();
            while (_groupDrawings.Count > 0)
            {
                Remove(_groupDrawings[0]);
            }
        }
    }
    /// <summary>
    /// Grouped shapes
    /// </summary>
    public class ExcelGroupShape : ExcelDrawing
    {
        internal ExcelGroupShape(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent = null, DrawingsCollectionType DrawingsType = DrawingsCollectionType.Worksheet) :
            base(drawings, node, NamespacePrefixes[(int)DrawingsType] + ":grpSp", NamespacePrefixes[(int)DrawingsType] + ":nvGrpSpPr/" + NamespacePrefixes[(int)DrawingsType] + ":cNvPr", parent, DrawingsType)
        {
            if (DrawingsType == DrawingsCollectionType.Chart)
            {
                node.OwnerDocument.DocumentElement.SetAttribute("xmlns:cdr", ExcelPackage.schemaChartDrawing);
                node.OwnerDocument.DocumentElement.SetAttribute("xmlns:a", ExcelPackage.schemaDrawings);
            }
            var grpNode = CreateNode(_topPath);
            if (grpNode.InnerXml == "")
            {
                Id = drawings._nextDrawingId++;
                grpNode.InnerXml = "<" + NamespacePrefixes[_prefixIndex] + ":nvGrpSpPr><" + NamespacePrefixes[_prefixIndex] + ":cNvPr name=\"\" id=\"" + Id + "\"><a:extLst><a:ext uri=\"{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}\"><a16:creationId id=\"{F33F4CE3-706D-4DC2-82DA-B596E3C8ACD0}\" xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\"/></a:ext></a:extLst></" + NamespacePrefixes[_prefixIndex] + ":cNvPr><" + NamespacePrefixes[_prefixIndex] + ":cNvGrpSpPr/></" + NamespacePrefixes[_prefixIndex] + ":nvGrpSpPr><" + NamespacePrefixes[_prefixIndex] + ":grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></" + NamespacePrefixes[_prefixIndex] + ":grpSpPr>";
            }

            switch (DrawingsType)
            {
                case DrawingsCollectionType.Chart:
                    if (parent == null)
                    {
                        int x = (int)(_drawings._screenWidth * EMU_PER_PIXEL * (From.X));
                        int y = (int)(_drawings._screenHeight * EMU_PER_PIXEL * (From.Y));
                        int cx, cy;
                        if (To == null)
                        {
                            cx = (int)(Size.Width / EMU_PER_PIXEL);
                            cy = (int)(Size.Height / EMU_PER_PIXEL);
                        }
                        else
                        {
                            cx = (int)(_drawings._screenWidth * EMU_PER_PIXEL * (To.X - From.X));
                            cy = (int)(_drawings._screenHeight * EMU_PER_PIXEL * (To.Y - From.Y));
                        }

                        XmlElement xFrmNode = GetXfrmNode(grpNode);
                        if (xFrmNode.ChildNodes.Count == 0)
                        {
                            CreateNode(xFrmNode, "a:off");
                            CreateNode(xFrmNode, "a:ext");
                            CreateNode(xFrmNode, "a:chOff");
                            CreateNode(xFrmNode, "a:chExt");
                        }
                        var offNode = (XmlElement)xFrmNode.SelectSingleNode("a:off", NameSpaceManager);
                        offNode.SetAttribute("x", x.ToString());
                        offNode.SetAttribute("y", y.ToString());
                        var extNode = (XmlElement)xFrmNode.SelectSingleNode("a:ext", NameSpaceManager);
                        extNode.SetAttribute("cx", cx.ToString());
                        extNode.SetAttribute("cy", cy.ToString());
                        Position = new ExcelDrawingCoordinate(drawings.NameSpaceManager, offNode, GetPositionSize);
                        Size = new ExcelDrawingSize(drawings.NameSpaceManager, extNode, GetPositionSize);
                        var chOffNode = (XmlElement)xFrmNode.SelectSingleNode("a:chOff", NameSpaceManager);
                        chOffNode.SetAttribute("x", x.ToString());
                        chOffNode.SetAttribute("y", y.ToString());
                        var chExtNode = (XmlElement)xFrmNode.SelectSingleNode("a:chExt", NameSpaceManager);
                        chExtNode.SetAttribute("cx", cx.ToString());
                        chExtNode.SetAttribute("cy", cy.ToString());
                    }
                    else
                    {
                        XmlElement xFrmNode = GetXfrmNode(grpNode);
                        var offNode = (XmlElement)xFrmNode.SelectSingleNode("a:off", NameSpaceManager);
                        var extNode = (XmlElement)xFrmNode.SelectSingleNode("a:ext", NameSpaceManager);
                        Position = new ExcelDrawingCoordinate(drawings.NameSpaceManager, offNode, GetPositionSize);
                        Size = new ExcelDrawingSize(drawings.NameSpaceManager, extNode, GetPositionSize);
                    }
                    break;
                case DrawingsCollectionType.Worksheet:
                default:
                    if (parent == null && node.SelectSingleNode("xdr:clientData", NameSpaceManager) == null)
                    {
                        node.AppendChild(grpNode.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings));
                    }
                    break;
            }
        }
        ExcelDrawingsGroup _groupDrawings = null;
        /// <summary>
        /// A collection of shapes
        /// </summary>
        public ExcelDrawingsGroup Drawings
        {
            get
            {
                if (_groupDrawings == null)
                {
                    if (string.IsNullOrEmpty(_topPath))
                    {
                        _groupDrawings = new ExcelDrawingsGroup(this, NameSpaceManager, TopNode, _collectionType);
                    }
                    else
                    {
                        if (_collectionType == DrawingsCollectionType.Chart)
                        {
                            _groupDrawings = new ExcelDrawingsGroup(this, NameSpaceManager, GetNode(_topPath), _collectionType);
                        }
                        else
                        {
                            _groupDrawings = new ExcelDrawingsGroup(this, NameSpaceManager, GetNode(_topPath));
                        }
                    }
                }
                return _groupDrawings;
            }
        }

        internal static void Validate(ExcelDrawing d, ExcelDrawings drawings, ExcelGroupShape grp)
        {
            if (d._drawings != drawings)
            {
                throw new InvalidOperationException("All drawings must be in the same worksheet.");
            }
            if (d._parent != null && d._parent != grp)
            {
                throw new InvalidOperationException($"The drawing {d.Name} is already in a group different from the other drawings.");
            }
            if (d._drawings._collectionType != drawings._collectionType)
            {
                throw new InvalidOperationException("Drawings need to be inside the same drawings type collection.");
            }
        }
        internal void SetPositionAndSizeFromChildren()
        {
            if (_collectionType == DrawingsCollectionType.Worksheet)
            {
                var pd = Drawings[0];
                pd.GetPositionSize();
                double t = pd._top, l = pd._left, b = pd._top + pd._height, r = pd._left + pd._width;
                for (int i = 1; i < Drawings.Count; i++)
                {
                    var d = Drawings[i];
                    d.GetPositionSize();
                    if (t > d._top)
                    {
                        t = d._top;
                    }
                    if (l > d._left)
                    {
                        l = d._left;
                    }
                    if (r < d._left + d._width)
                    {
                        r = d._left + d._width;
                    }
                    if (b < d._top + d._height)
                    {
                        b = d._top + d._height;
                    }
                }
                SetPosition((int)t, (int)l, false);
                SetSize((int)(r - l), (int)(b - t));
            }
            else if (_collectionType == DrawingsCollectionType.Chart)
            {
                long l, t, r, b;
                GetDrawingBoundries(out l, out t, out r, out b);

                foreach (var d in Drawings)
                {
                    GetDrawingBoundries(out long dl, out long dt, out long dr, out long db);
                    l = Math.Min(l, dl);
                    t = Math.Min(t, dt);
                    r = Math.Max(r, dr);
                    b = Math.Max(b, db);
                }
                long w = r - l;
                long h = b - t;
                Size.Width = w;
                Size.Height = h;
                Position.X = (int)l;
                Position.Y = (int)t;
                xFrmPosition.X = (int)l;
                xFrmPosition.Y = (int)t;
                xFrmSize.Width = w;
                xFrmSize.Height = h;
                xFrmChildPosition.X = (int)l;
                xFrmChildPosition.Y = (int)t;
                xFrmChildSize.Width = w;
                xFrmChildSize.Height = h;
                var off = TopNode.SelectSingleNode("cdr:grpSp/cdr:grpSpPr/a:xfrm/a:off", NameSpaceManager);
                off.Attributes["x"].Value = l.ToString();
                off.Attributes["y"].Value = t.ToString();
                var ext = TopNode.SelectSingleNode("cdr:grpSp/cdr:grpSpPr/a:xfrm/a:ext", NameSpaceManager);
                ext.Attributes["cx"].Value = w.ToString();
                ext.Attributes["cy"].Value = h.ToString();
                var chOff = TopNode.SelectSingleNode("cdr:grpSp/cdr:grpSpPr/a:xfrm/a:chOff", NameSpaceManager);
                chOff.Attributes["x"].Value = l.ToString();
                chOff.Attributes["y"].Value = t.ToString();
                var chExt = TopNode.SelectSingleNode("cdr:grpSp/cdr:grpSpPr/a:xfrm/a:chExt", NameSpaceManager);
                chExt.Attributes["cx"].Value = w.ToString();
                chExt.Attributes["cy"].Value = h.ToString();

                From.X = l / (_drawings._screenWidth * EMU_PER_PIXEL);
                From.Y = t / (_drawings._screenHeight * EMU_PER_PIXEL);
                To.X = r / (_drawings._screenWidth * EMU_PER_PIXEL);
                To.Y = b / (_drawings._screenHeight * EMU_PER_PIXEL);
                From.UpdateXml();
                To.UpdateXml();
            }

        }

        private void GetDrawingBoundries(out long l, out long t, out long r, out long b)
        {
            if (Drawings[0]._frmXPosition == null)
            {
                l = Drawings[0].GetPixelLeft() * EMU_PER_PIXEL;
                t = Drawings[0].GetPixelTop() * EMU_PER_PIXEL;
                r = l + ((long)Drawings[0].GetPixelWidth() * EMU_PER_PIXEL);
                b = t + ((long)Drawings[0].GetPixelHeight() * EMU_PER_PIXEL);
            }
            else
            {
                l = Drawings[0]._frmXPosition.X;
                t = Drawings[0]._frmXPosition.Y;
                r = Drawings[0]._frmXPosition.X + Drawings[0]._frmXSize.Width;
                b = Drawings[0]._frmXPosition.Y + Drawings[0]._frmXSize.Height;
            }
        }

        internal void AdjustChildrenForResizeRow(double prevTop)
        {
            var top = GetPixelTop();
            var diff = top - prevTop;
            if (diff != 0)
            {
                for (int i = 0; i < Drawings.Count; i++)
                {
                    Drawings[i].SetPixelTop(Drawings[i]._top + diff);
                    Drawings[i].Position.UpdateXml();
                }
            }
        }
        internal void AdjustChildrenForResizeColumn(double prevLeft)
        {
            var left = GetPixelLeft();
            var diff = left - prevLeft;
            if (diff != 0)
            {
                for (int i = 0; i < Drawings.Count; i++)
                {
                    Drawings[i].SetPixelLeft(Drawings[i]._left + diff);
                    Drawings[i].Position.UpdateXml();
                }
            }
        }

        ExcelDrawingCoordinate _xFrmPosition = null;
        internal ExcelDrawingCoordinate xFrmPosition
        {
            get
            {
                if (_xFrmPosition == null)
                {
                    _xFrmPosition = new ExcelDrawingCoordinate(NameSpaceManager, GetNode(NamespacePrefixes[_prefixIndex] + ":grpSp/" + NamespacePrefixes[_prefixIndex] + ":grpSpPr/a:xfrm/a:off"));
                }
                return _xFrmPosition;
            }
        }
        ExcelDrawingSize _xFrmSize = null;
        internal ExcelDrawingSize xFrmSize
        {
            get
            {
                if (_xFrmSize == null)
                {
                    _xFrmSize = new ExcelDrawingSize(NameSpaceManager, GetNode(NamespacePrefixes[_prefixIndex] + ":grpSp/" + NamespacePrefixes[_prefixIndex] + ":grpSpPr/a:xfrm/a:ext"));
                }
                return _xFrmSize;
            }
        }
        ExcelDrawingCoordinate _xFrmChildPosition = null;
        internal ExcelDrawingCoordinate xFrmChildPosition
        {
            get
            {
                if (_xFrmChildPosition == null)
                {
                    _xFrmChildPosition = new ExcelDrawingCoordinate(NameSpaceManager, GetNode(NamespacePrefixes[_prefixIndex] + ":grpSp/" + NamespacePrefixes[_prefixIndex] + ":grpSpPr/a:xfrm/a:chOff"));
                }
                return _xFrmChildPosition;
            }
        }
        ExcelDrawingSize _xFrmChildSize = null;
        internal ExcelDrawingSize xFrmChildSize
        {
            get
            {
                if (_xFrmChildSize == null)
                {
                    _xFrmChildSize = new ExcelDrawingSize(NameSpaceManager, GetNode(NamespacePrefixes[_prefixIndex] + ":grpSp/" + NamespacePrefixes[_prefixIndex] + ":grpSpPr/a:xfrm/a:chExt"));
                }
                return _xFrmChildSize;
            }
        }
        /// <summary>
        /// The type of drawing
        /// </summary>
        public override eDrawingType DrawingType
        {
            get
            {
                return eDrawingType.GroupShape;
            }
        }

        internal override void SaveDrawing(bool hasLoadedPivotTables)
        {
            base.SaveDrawing(hasLoadedPivotTables);

            foreach (var d in Drawings)
            {
                d.SaveDrawing(hasLoadedPivotTables);
            }
        }
    }
}