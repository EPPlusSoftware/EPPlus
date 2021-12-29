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
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Fill properties for drawing objects like lines etc, that don't have blip- and pattern- fills
    /// </summary>
    public class ExcelDrawingFillBasic : XmlHelper, IDisposable
    {
        /// <summary>
        /// XPath
        /// </summary>
        internal protected string _fillPath;
        /// <summary>
        /// The fill xml element
        /// </summary>
        internal protected XmlNode _fillNode;
        /// <summary>
        /// The drawings collection
        /// </summary>
        internal protected ExcelDrawing _drawing;
        /// <summary>
        /// The fill type node.
        /// </summary>
        internal protected XmlNode _fillTypeNode = null;
        internal Action _initXml;
        internal ExcelDrawingFillBasic(ExcelPackage pck, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string fillPath, string[] schemaNodeOrderBefore, bool doLoad, Action initXml = null) :
            base(nameSpaceManager, topNode)
        {
            AddSchemaNodeOrder(schemaNodeOrderBefore, new string[] { "xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "blipFill", "gradFill", "noFill", "pattFill", "grpFill", "ln", "effectLst", "effectDag", "highlight", "latin", "cs", "sym", "ea", "hlinkClick", "hlinkMouseOver", "rtl" });
            _fillPath = fillPath;
            _initXml = initXml;
            SetFillNodes(topNode);
            //Setfill node
            if (doLoad && _fillNode != null)
            {
                LoadFill();
            }
            if (pck != null)
            {
                pck.BeforeSave.Add(BeforeSave);
            }
        }
        internal void SetTopNode(XmlNode topNode)
        {
            TopNode = topNode;
            SetFillNodes(topNode);
            _fillTypeNode = null;
            LoadFill();
        }
        private void SetFillNodes(XmlNode topNode)
        {
            if (string.IsNullOrEmpty(_fillPath))
            {
                _fillNode = topNode;
                if (topNode.LocalName.EndsWith("Fill"))  //Theme nodes will have the fillnode as topnode
                {
                    _fillTypeNode = _fillNode;
                }
            }
            else
            {
                _fillNode = topNode.SelectSingleNode(_fillPath, NameSpaceManager);
            }
        }

        internal virtual void BeforeSave()
        {
            if(_gradientFill!=null)
            {
                _gradientFill.UpdateXml();
            }
        }
        /// <summary>
        /// Loads the fill from xml
        /// </summary>
        internal protected virtual void LoadFill()
        {
            if (_fillTypeNode == null) _fillTypeNode = _fillNode.SelectSingleNode("a:solidFill", NameSpaceManager);
            if (_fillTypeNode == null) _fillTypeNode = _fillNode.SelectSingleNode("a:gradFill", NameSpaceManager);
            if (_fillTypeNode == null) _fillTypeNode = _fillNode.SelectSingleNode("a:noFill", NameSpaceManager);
            if (_fillTypeNode == null)
                return;

            switch (_fillTypeNode.LocalName)
            {
                case "solidFill":
                    _style = eFillStyle.SolidFill;
                    _solidFill = new ExcelDrawingSolidFill(NameSpaceManager, _fillTypeNode, "", SchemaNodeOrder, _initXml);
                    break;
                case "gradFill":
                    _style = eFillStyle.GradientFill;
                    _gradientFill = new ExcelDrawingGradientFill(NameSpaceManager, _fillTypeNode, SchemaNodeOrder, _initXml);
                    break;
                default:
                    _style = eFillStyle.NoFill;
                    break;
            }
        }
        internal void SetFromXml(ExcelDrawingFill fill)
        {
            Style = fill.Style;
            var copyFromFillElement = (XmlElement)fill._fillTypeNode;
            foreach (XmlAttribute a in copyFromFillElement.Attributes)
            {
                ((XmlElement)_fillTypeNode).SetAttribute(a.Name, a.NamespaceURI, a.Value);
            }
            _fillTypeNode.InnerXml = copyFromFillElement.InnerXml;
            if(fill.Style==eFillStyle.BlipFill)
            {
                XmlAttribute relAttr=(XmlAttribute)_fillTypeNode.SelectSingleNode("a:blip/@r:embed", NameSpaceManager);
                if(relAttr?.Value!=null)
                {
                    relAttr.OwnerElement.Attributes.Remove(relAttr);
                }
            }
            LoadFill();
            if(Style==eFillStyle.BlipFill)
            {
                
                ((ExcelDrawingFill)this).BlipFill.Image = fill.BlipFill.Image;

            }
        }

        private void CreateImageRelation(ExcelDrawingFill fill, XmlElement copyFromFillElement)
        {
            IPictureContainer pic = fill.BlipFill;

        }

        internal string GetFromXml()
        {
            return _fillTypeNode.OuterXml;
        }
        internal virtual void SetFillProperty()
        {   
            if (_fillNode==null)
            {
                InitSpPr(eFillStyle.SolidFill);
                Style = eFillStyle.SolidFill;   //This will create the _fillNode
                _solidFill = new ExcelDrawingSolidFill(NameSpaceManager, _fillTypeNode, "", SchemaNodeOrder, _initXml);
                return; 
            }

            _solidFill = null;
            _gradientFill = null;

            switch (_fillTypeNode.LocalName)
            {
                case "solidFill":
                    _solidFill = new ExcelDrawingSolidFill(NameSpaceManager, _fillTypeNode, "",SchemaNodeOrder, _initXml);
                    break;
                case "gradFill":
                    _gradientFill = new ExcelDrawingGradientFill(NameSpaceManager, _fillTypeNode, SchemaNodeOrder, _initXml);
                    break;
                default:
                    if(this is ExcelDrawingFillBasic && _style!=eFillStyle.NoFill)
                    {
                        throw new ArgumentException("Style", $"Style {Style} cannot be applied to this object.");
                    }
                    break;
            }
        }
        bool isSpInit = false;
        private void InitSpPr(eFillStyle style)
        {
            if (isSpInit == false)
            {
                if (!string.IsNullOrEmpty(_fillPath) && !ExistsNode(_fillPath) && CreateNodeUntil(_fillPath, "spPr", out XmlNode spPrNode))
                {
                    if(_fillPath.EndsWith("ln"))
                    {
                        spPrNode.InnerXml = $"<a:ln><a:noFill/></a:ln ><a:effectLst/><a:sp3d/>";
                    }
                    else
                    {
                        spPrNode.InnerXml = $"<a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/><a:sp3d/>";
                    }
                    _fillNode = GetNode(_fillPath);
                    _fillTypeNode = _fillNode.FirstChild;
                }
                else if(_fillTypeNode==null)
                {
                    if(_fillNode==null)
                    {
                        _fillNode = GetNode(_fillPath);
                    }
                    if (!_fillNode.HasChildNodes)
                    {
                        _fillNode.InnerXml = $"<a:{GetStyleText(style)}/>";
                    }
                    LoadFill();
                }
            }
            isSpInit = true;
        }
        internal eFillStyle? _style=null;
        /// <summary>
        /// Fill style
        /// </summary>
        public eFillStyle Style
        {
            get
            {
                return _style??eFillStyle.NoFill;
            }
            set
            {
                if (_style == value) return;
                _initXml?.Invoke();
                if (value == eFillStyle.GroupFill)
                {
                    throw new NotImplementedException("Fillstyle not implemented");
                }
                else
                {
                    _style = value;
                    InitSpPr(value);
                    CreateFillTopNode(value);
                    SetFillProperty();
                }
            }
        }
        const string ColorPath = "a:srgbClr/@val";
        /// <summary>
        /// Fill color for solid fills.
        /// Other fill styles will return Color.Empty.
        /// Setting this propery will set the Type to SolidFill with the specified color.
        /// </summary>
        public Color Color
        {
            get
            {
                if (Style != eFillStyle.SolidFill) return Color.Empty;
                if (SolidFill.Color.ColorType != eDrawingColorType.Rgb) return Color.Empty;
                var col = SolidFill.Color.RgbColor.Color;
                if(col == Color.Empty)
                {
                    return Color.FromArgb(79, 129, 189);
                }
                else
                {
                    return col;
                }
            }
            set
            {
                _initXml?.Invoke();
                Style = eFillStyle.SolidFill;
                SolidFill.Color.SetRgbColor(value);
            }
        }
        private ExcelDrawingSolidFill _solidFill =null;
        
        /// <summary>
        /// Reference solid fill properties
        /// This property is only accessable when Type is set to SolidFill
        /// </summary>
        public ExcelDrawingSolidFill SolidFill
        {
            get
            {
                if(Style==eFillStyle.SolidFill && _solidFill==null)
                {
                    _solidFill = new ExcelDrawingSolidFill(NameSpaceManager, _fillTypeNode, "", SchemaNodeOrder, _initXml);
                }
                return _solidFill;
            }
        }
        private ExcelDrawingGradientFill _gradientFill = null;
        /// <summary>
        /// Reference gradient fill properties
        /// This property is only accessable when Type is set to GradientFill
        /// </summary>
        public ExcelDrawingGradientFill GradientFill
        {
            get
            {
                return _gradientFill;
            }
        }
        /// <summary>
        /// Transparancy in percent from a solid fill. 
        /// This is the same as 100-Fill.Transform.Alpha
        /// </summary>
        public int Transparancy
        {
            get
            {
                if (_solidFill == null) return 0;
                return (int)(100-_solidFill.Color.Transforms.FindValue(eColorTransformType.Alpha));
            }
            set
            {
                if (_solidFill == null) throw new InvalidOperationException("Transparency can only be set when Type is set to SolidFill.");
                var alphaItem = _solidFill.Color.Transforms.Find(eColorTransformType.Alpha);
                if(alphaItem==null)
                {
                    _solidFill.Color.Transforms.AddAlpha(100 - value);
                }
                else
                {
                    alphaItem.Value = 100 - value;
                }
            }
        }
        private void CreateFillTopNode(eFillStyle value)
        {
            if (_fillNode == TopNode)
            {
                if(_fillNode== _fillTypeNode)
                {
                    var node=_fillTypeNode.OwnerDocument.CreateElement("a", GetStyleText(value), ExcelPackage.schemaDrawings);
                    _fillTypeNode.ParentNode.InsertBefore(node, _fillTypeNode);
                    _fillTypeNode.ParentNode.RemoveChild(_fillTypeNode);
                    _fillTypeNode = node;
                    _fillNode = node;
                    TopNode = node;
                }
                else
                {
                    _fillTypeNode = CreateNode("a:" + GetStyleText(value));
                }
            }
            else
            {
                if (_fillTypeNode != null)
                {
                    _fillTypeNode.ParentNode.RemoveChild(_fillTypeNode);
                }
                _fillTypeNode = CreateNode(_fillPath + "/a:" + GetStyleText(value), false);
                if(_fillNode==null)
                {
                    _fillNode = _fillTypeNode.ParentNode;
                }
            }
        }

        internal static eFillStyle GetStyleEnum(string name)
        {
            switch (name)
            {
                case "noFill":
                    return eFillStyle.NoFill;
                case "blipFill":
                    return eFillStyle.BlipFill;
                case "gradFill":
                    return eFillStyle.GradientFill;
                case "grpFill":
                    return eFillStyle.GroupFill;
                case "pattFill":
                    return eFillStyle.PatternFill;
                default:
                    return eFillStyle.SolidFill;
            }
        }

        internal static string GetStyleText(eFillStyle style)
        {
            switch (style)
            {
                case eFillStyle.BlipFill:
                    return "blipFill";
                case eFillStyle.GradientFill:
                    return "gradFill";
                case eFillStyle.GroupFill:
                    return "grpFill";
                case eFillStyle.NoFill:
                    return "noFill";
                case eFillStyle.PatternFill:
                    return "pattFill";
                default:
                    return "solidFill";
            }
        }

        /// <summary>
        /// Disposes the object
        /// </summary>
        public void Dispose()
        {
            _fillNode = null;
            _solidFill = null;
            _gradientFill = null;
        }
    }
}

