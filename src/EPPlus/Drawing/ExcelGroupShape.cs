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
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// A collection of sub drawings to a group drawing
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
            ExcelGroupShape.Validate(drawing, _parent._drawings);
            AppendDrawingNode(drawing.TopNode);
            drawing._parent = _parent;
            _groupDrawings.Add(drawing);
            _drawingNames.Add(drawing.Name, _groupDrawings.Count - 1);
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
                grpNode.InnerXml = "<xdr:nvGrpSpPr><xdr:cNvPr name=\"\" id=\"3\"><a:extLst><a:ext uri=\"{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}\"><a16:creationId id=\"{F33F4CE3-706D-4DC2-82DA-B596E3C8ACD0}\" xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off y=\"561975\" x=\"3028950\"/><a:ext cy=\"2524125\" cx=\"3152775\"/><a:chOff y=\"561975\" x=\"3028950\"/><a:chExt cy=\"2524125\" cx=\"3152775\"/></a:xfrm></xdr:grpSpPr>";
            }
            CreateNode("xdr:clientData");
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

        internal static void Validate(ExcelDrawing d, ExcelDrawings drawings)
        {
            if (d._drawings != drawings)
            {
                throw new InvalidOperationException("All drawings must be in the same worksheet.");
            }
            if (d._parent != null)
            {
                throw new InvalidOperationException($"The drawing {d.Name} is already in a group."); ;
            }
        }
        internal void SetPositionAndSizeFromChildren()
        {
            double t = Drawings[0]._top, l = Drawings[0]._left, h = Drawings[0]._top + Drawings[0]._height, w=Drawings[0]._left + Drawings[0]._width;
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
                if (w < Drawings[i]._left+Drawings[i]._width)
                {
                    w = Drawings[i]._left + Drawings[i]._width;
                }
                if (t < Drawings[i]._top + Drawings[i]._height)
                {
                    h = Drawings[i]._top + Drawings[i]._height;
                }
            }

            SetPosition((int)t, (int)l);
            SetSize((int)(w - l), (int)(h - t));
        }
    }
}