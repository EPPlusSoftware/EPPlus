using EPPlus.DrawingRenderer.Utils;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace EPPlus.DrawingRenderer.ShapeDefinitions
{
    public class DrawingPath
    {
        public DrawingPath(DrawingPath clone)
        {
            Width = clone.Width;
            Height = clone.Height;
            Fill = clone.Fill;
            Stroke = clone.Stroke;
            ExtrusionOk = clone.ExtrusionOk;
            foreach (var p in clone.Paths)
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
                else if (xr.LocalName == "path" && xr.NodeType == XmlNodeType.EndElement)
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

        public DrawingPath Clone() => new DrawingPath(this);

        public bool Stroke { get; set; }
        public bool ExtrusionOk { get; set; }
        public PathFillMode Fill { get; set; }
        public double? Width { get; set; }
        public double? Height { get; set; }
        public List<PathsBase> Paths { get; set; } = new List<PathsBase>();
    }
}
