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
                ExcelDrawing grpDraw;
                switch (node.LocalName)
                {
                    case "sp":
                        grpDraw = new ExcelShape(_parent._drawings, node, _parent);
                        break;
                    case "pic":
                        grpDraw=new ExcelPicture(_parent._drawings, node, _parent);
                        break;
                    case "graphicFrame":
                        grpDraw = ExcelChart.GetChart(_parent._drawings, node, _parent);
                        break;
                    case "grpSp":
                        grpDraw = new ExcelGroupShape(_parent._drawings, node, _parent);
                        break;
                    case "cxnSp":
                        grpDraw = new ExcelConnectionShape(_parent._drawings, node, _parent);
                        break;
                    default:
                        continue;
                }
                _groupDrawings.Add(grpDraw);
                _drawingNames.Add(grpDraw.Name, _groupDrawings.Count-1);
            }
        }
        /// <summary>
        /// Adds a drawing to the group
        /// </summary>
        /// <param name="drawing"></param>
        public void Add(ExcelDrawing drawing)
        {
            ExcelGroupShape.Validate(drawing, _parent._drawings);
            drawing._parent = _parent;
            _groupDrawings.Add(drawing);
            _drawingNames.Add(drawing.Name, _groupDrawings.Count - 1);
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
    }
}