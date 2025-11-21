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
using OfficeOpenXml.Utils.XML;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace EPPlusImageRenderer.ShapeDefinitions
{
    internal partial class PresetShapeDefinitions
    {
        public static void LoadPresetShapeDefinitionFromXml()
        {
            var xmlFile = Directory.GetCurrentDirectory() + "\\resource\\presetShapeDefinitions.xml";

            try
            {
                var ms = new MemoryStream(File.ReadAllBytes(xmlFile));
#if NET35
                var xr = XmlReader.Create(ms, new XmlReaderSettings()
                {
                    ProhibitDtd=true,
                    IgnoreWhitespace = true
                });
#else
                var xr = XmlReader.Create(ms, new XmlReaderSettings()
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    IgnoreWhitespace = true,
                    Async = true
                });
#endif

                while (xr.Read())
                {
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        if (xr.LocalName != "presetShapeDefinitons")
                        {
                            var item = LoadPresetShapeDefinition(xr);
                            if (item.Style == eShapeStyle.Sun)
                            {
                                item.translateCoordinate = new Coordinate(0, 0);
                            }
                            _shapeDefinitions.Add(item.Style, item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw (new IOException("Cannot preset shape definitions file:presetShapeDefinitions.xml", ex));
            }
        }

        public static bool LoadPresetShapeDefinitionFromXmlTriangle()
        {
            _shapeDefinitions = new Dictionary<eShapeStyle, ShapeDefinition>();

            var xmlFile = Directory.GetCurrentDirectory() + "\\resource\\triangleOnly.xml";
            var fs = new FileStream(xmlFile, FileMode.Open, FileAccess.Read);
#if NET35
                var xr = XmlReader.Create(fs, new XmlReaderSettings()
                {
                    ProhibitDtd=true,
                    IgnoreWhitespace = true
                });
#else
            var xr = XmlReader.Create(fs, new XmlReaderSettings()
            {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreWhitespace = true,
                Async = true
            });
#endif
            while (xr.Read())
            {
                if (xr.NodeType == XmlNodeType.Element)
                {
                    if (xr.LocalName == "triangle" && xr.NodeType != XmlNodeType.EndElement)
                    {
                        var item = LoadPresetShapeDefinition(xr);
                        _shapeDefinitions.Add(item.Style, item);
                    }
                }
            }
            return true;
        }

        private static ShapeDefinition LoadPresetShapeDefinition(XmlReader xr)
        {
            var style = xr.LocalName.ToEnum<eShapeStyle>();
            if (!style.HasValue) throw new InvalidOperationException();
            var psd = new ShapeDefinition
            {
                Style = style.Value,
            };
            LoadFromXml(psd, xr);
            return psd;
        }

        private static void LoadFromXml(ShapeDefinition psd, XmlReader xr)
        {
            while (xr.Read())
            {
                if (xr.NodeType == XmlNodeType.Element)
                {
                    switch (xr.LocalName)
                    {
                        case "avLst":
                            psd.ShapeAdjustValues = LoadShapeGuides(xr);
                            break;
                        case "gdLst":
                            psd.ShapeGuides = LoadShapeGuides(xr);
                            break;
                        case "ahLst":
                            psd.ShapeAdjustHandles = LoadAdjustHandle(xr);
                            break;
                        case "cxnLst":
                            psd.ShapeConnectionSite = LoadConnectionLst(xr);
                            break;
                        case "rect":
                            psd.TextBoxRect = LoadRect(xr);
                            break;
                        case "pathLst":
                            psd.ShapePaths = LoadShapePaths(xr);
                            break;
                    }
                }
                else if (xr.NodeType == XmlNodeType.EndElement && xr.LocalName.Equals(psd.Style.ToString(), StringComparison.InvariantCultureIgnoreCase))
                {
                    return;
                }
            }
        }

        private static TextBoxRect LoadRect(XmlReader xr)
        {
           //  < rect l = "l" t = "y1" r = "x4" b = "b" xmlns = "http://schemas.openxmlformats.org/drawingml/2006/main" />
            var rect = new TextBoxRect()
            {
                TopName = xr.GetAttribute("t"),
                BottomName = xr.GetAttribute("b"),
                LeftName = xr.GetAttribute("l"),
                RightName = xr.GetAttribute("r")
            };

            return rect;
        }

        private static List<ShapeConnectionSite> LoadConnectionLst(XmlReader xr)
        {
            var shapeConnectionSite = new List<ShapeConnectionSite>();

            while (xr.Read() && (xr.NodeType != XmlNodeType.EndElement && xr.LocalName != "cxnLst"))
            {
                var newConnection = new ShapeConnectionSite();

                var attrStr = xr.GetAttribute("ang");

                newConnection.Angle = xr.GetAttribute("ang");
                xr.Read();

                newConnection.PositionCoordinate = new ShapePositionCoordinate() { X = xr.GetAttribute("x"), Y = xr.GetAttribute("y") };

                shapeConnectionSite.Add(newConnection);
                xr.Read();

                if (xr.LocalName == "cxnLst" && xr.NodeType == XmlNodeType.EndElement)
                {
                    break;
                }
            }

            return shapeConnectionSite;
        }

        private static List<ShapeAdjustHandleBase> LoadAdjustHandle(XmlReader xr)
        {
            var l = new List<ShapeAdjustHandleBase>();
            var name = xr.LocalName;
            while (xr.Read())
            {
                if (xr.NodeType == XmlNodeType.Element)
                {
                    switch (xr.LocalName)
                    {
                        case "ahXY":
                            l.Add(ShapeAdjustHandleBase.CreateXy(xr));
                            break;
                        case "ahPolar":
                            l.Add(ShapeAdjustHandleBase.CreatePolar(xr));
                            break;
                        case "pos":
                            l[l.Count - 1].PositionCoordinate = new ShapePositionCoordinate() { X = xr.GetAttribute("x"), Y = xr.GetAttribute("y") };
                            break;
                    }
                }
                else if (xr.NodeType == XmlNodeType.EndElement && xr.LocalName == name)
                {
                    break;
                }
            }
            return l;
        }

        private static List<ShapeGuide> LoadShapeGuides(XmlReader xr)
        {
            var l = new List<ShapeGuide>();
            var name = xr.LocalName;
            while (xr.Read())
            {
                if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "gd")
                {
                    l.Add(new ShapeGuide() { Name = xr.GetAttribute("name"), Formula = xr.GetAttribute("fmla") });
                }

                if (xr.NodeType == XmlNodeType.EndElement && xr.LocalName == name)
                {
                    break;
                }
            }
            return l;
        }


        private static List<DrawingPath> LoadShapePaths(XmlReader xr)
        {
            var list = new List<DrawingPath>();
            while (xr.Read())
            {
                if (xr.LocalName == "path" && xr.NodeType == XmlNodeType.Element)
                {
                    list.Add(new DrawingPath(xr));
                }
                else if (xr.IsEndElementWithName("pathLst"))
                {
                    break;
                }
            }
            return list;
        }
    }
}