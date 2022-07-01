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
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

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
        XmlNamespaceManager _nsm;
        XmlNode _topNode;
        internal ExcelDrawingsGroup(ExcelGroupShape parent, XmlNamespaceManager nsm, XmlNode topNode)
        {
            _parent = parent;
            _nsm = nsm;
            _topNode = topNode;
            _drawingNames = new Dictionary<string, int>();
            AddDrawings();
        }
        private void AddDrawings()
        {
            _groupDrawings = new List<ExcelDrawing>();
            foreach (XmlNode node in _topNode.ChildNodes)
            {
                
                if (node.LocalName != "nvGrpSpPr" && node.LocalName != "grpSpPr")
                {
                    var grpDraw = ExcelDrawing.GetDrawingFromNode(_parent._drawings, node, (XmlElement)node, _parent);
                    _groupDrawings.Add(grpDraw);
                    _drawingNames.Add(grpDraw.Name, _groupDrawings.Count - 1);
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
            XmlElement xFrmNode = d.GetFrmxNode(node);
            if (xFrmNode.ChildNodes.Count == 0)
            {
                d.CreateNode(xFrmNode, "a:off");
                d.CreateNode(xFrmNode, "a:ext");
            }
            var offNode = (XmlElement)xFrmNode.SelectSingleNode("a:off", _nsm);
            offNode.SetAttribute("y", (top * ExcelDrawing.EMU_PER_PIXEL).ToString());
            offNode.SetAttribute("x", (left * ExcelDrawing.EMU_PER_PIXEL).ToString());
            var extNode = (XmlElement)xFrmNode.SelectSingleNode("a:ext",_nsm);
            extNode.SetAttribute("cy", Math.Round(height * ExcelDrawing.EMU_PER_PIXEL, 0).ToString());
            extNode.SetAttribute("cx", Math.Round(width * ExcelDrawing.EMU_PER_PIXEL, 0).ToString());
            
            d.SetGroupChild(offNode, extNode);
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
                //drawingNode = xmlDoc.CreateElement("mc", "AlternateContent", ExcelPackage.schemaMarkupCompatibility);
                drawingNode = _parent.TopNode.ParentNode.ParentNode.CloneNode(false);
                var choiceNode= _parent.TopNode.ParentNode.CloneNode(false);
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
            d.SetPosition(top, left);
            d.SetSize((int)width, (int)height);
        }

        private XmlNode CreateAnchorNode(XmlNode drawingNode)
        {
            var topNode = _parent.TopNode.CloneNode(false);
            topNode.AppendChild(_parent.TopNode.GetChildAtPosition(0).CloneNode(true));
            topNode.AppendChild(_parent.TopNode.GetChildAtPosition(1).CloneNode(true));
            topNode.AppendChild(drawingNode);
            var ix = 3;
            while(ix< _parent.TopNode.ChildNodes.Count)
            {
                topNode.AppendChild(_parent.TopNode.ChildNodes[ix].CloneNode(true));
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
                    return _groupDrawings[_parent._drawings._drawingNames[Name]];
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
            _parent._drawings.ReIndexNames(ix, 1);
            _parent._drawings._drawingNames.Add(drawing.Name, ix);

            //Remove 
            if (_parent.Drawings.Count == 0)
            {
                _parent._drawings.Remove(_parent);
            }
            drawing._parent = null;
        }
        /// <summary>
        /// Removes all children drawings from the group.
        /// </summary>
        public void Clear()
        {
            CheckNotDisposed();
            while (_groupDrawings.Count>0)
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
        internal ExcelGroupShape(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent = null) : 
            base(drawings, node, "xdr:grpSp", "xdr:nvGrpSpPr/xdr:cNvPr", parent)
        {
            var grpNode = CreateNode(_topPath);
            if (grpNode.InnerXml == "")
            {
                grpNode.InnerXml = "<xdr:nvGrpSpPr><xdr:cNvPr name=\"\" id=\"3\"><a:extLst><a:ext uri=\"{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}\"><a16:creationId id=\"{F33F4CE3-706D-4DC2-82DA-B596E3C8ACD0}\" xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off y=\"0\" x=\"0\"/><a:ext cy=\"0\" cx=\"0\"/><a:chOff y=\"0\" x=\"0\"/><a:chExt cy=\"0\" cx=\"0\"/></a:xfrm></xdr:grpSpPr>";
            }
            if(parent==null) CreateNode("xdr:clientData");
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
                        _groupDrawings = new ExcelDrawingsGroup(this, NameSpaceManager, TopNode);
                    }
                    else
                    {
                        _groupDrawings = new ExcelDrawingsGroup(this, NameSpaceManager, GetNode(_topPath));
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
            if (d._parent != null && d._parent!=grp)
            {
                throw new InvalidOperationException($"The drawing {d.Name} is already in a group different from the other drawings."); ;
            }
        }
        internal void SetPositionAndSizeFromChildren()
        {
            double t = Drawings[0]._top, l = Drawings[0]._left, b = Drawings[0]._top + Drawings[0]._height, r=Drawings[0]._left + Drawings[0]._width;
            for(int i=1;i<Drawings.Count;i++)
            {
                if(t>Drawings[i]._top)
                {
                   t = Drawings[i]._top;
                }
                if (l > Drawings[i]._left)
                {
                    l = Drawings[i]._left;
                }
                if (r < Drawings[i]._left+Drawings[i]._width)
                {
                    r = Drawings[i]._left + Drawings[i]._width;
                }
                if (b < Drawings[i]._top + Drawings[i]._height)
                {
                    b = Drawings[i]._top + Drawings[i]._height;
                }
            }

            SetPosition((int)t, (int)l, false);
            SetSize((int)(r - l), (int)(b - t));

            SetxFrmPosition();
        }

        private void SetxFrmPosition()
        {
            xFrmPosition.X = (int)(_left * EMU_PER_PIXEL);
            xFrmPosition.Y = (int)(_top * EMU_PER_PIXEL);
            xFrmSize.Width = (long)(_width * EMU_PER_PIXEL);
            xFrmSize.Height = (long)(_height * EMU_PER_PIXEL);

            xFrmChildPosition.X = (int)(_left * EMU_PER_PIXEL);
            xFrmChildPosition.Y = (int)(_top * EMU_PER_PIXEL);
            xFrmChildSize.Width = (long)(_width * EMU_PER_PIXEL);
            xFrmChildSize.Height = (long)(_height * EMU_PER_PIXEL);
        }

        ExcelDrawingCoordinate _xFrmPosition = null;
        internal ExcelDrawingCoordinate xFrmPosition
        {
            get
            {
                if(_xFrmPosition==null)
                {
                    _xFrmPosition= new ExcelDrawingCoordinate(NameSpaceManager, GetNode("xdr:grpSp/xdr:grpSpPr/a:xfrm/a:off")); 
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
                    _xFrmSize = new ExcelDrawingSize(NameSpaceManager, GetNode("xdr:grpSp/xdr:grpSpPr/a:xfrm/a:ext")); 
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
                    _xFrmChildPosition = new ExcelDrawingCoordinate(NameSpaceManager, GetNode("xdr:grpSp/xdr:grpSpPr/a:xfrm/a:chOff"));
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
                    _xFrmChildSize = new ExcelDrawingSize(NameSpaceManager, GetNode("xdr:grpSp/xdr:grpSpPr/a:xfrm/a:chExt"));
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
    }
    
}