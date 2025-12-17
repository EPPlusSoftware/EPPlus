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
using System;
using System.Collections.Generic;
using System.IO;

#if !NET35
//using System.Reflection.Metadata.Ecma335;
using System.Threading.Tasks;
//using System.Windows.Markup;
#endif
using System.Xml;

namespace EPPlusImageRenderer.ShapeDefinitions
{
    internal partial class PresetShapeDefinitions
    {
        static object _syncRoot=new object();
        static Dictionary<eShapeStyle, ShapeDefinition> _shapeDefinitions=null;
        public static Dictionary<eShapeStyle, ShapeDefinition> ShapeDefinitions 
        {
            get
            {
                lock (_syncRoot)
                {
                    if (_shapeDefinitions == null)
                    {
                        _shapeDefinitions = new Dictionary<eShapeStyle, ShapeDefinition>();
#if NET35
                        LoadPresetShapeDefinitionFromXml();
#else
                        Task.Run(() => LoadPresetShapeDefinitionFromXmlAsync()).Wait();
#endif

                    }
                    return _shapeDefinitions;
                }
            }
        }

#if !NET35
        public static async Task LoadPresetShapeDefinitionFromXmlAsync()
        {
            var xmlFile = Directory.GetCurrentDirectory() + "\\resource\\presetShapeDefinitions.xml";

            try
            {
                var ms=new MemoryStream(File.ReadAllBytes(xmlFile));
                var xr = XmlReader.Create(ms, new XmlReaderSettings()
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    IgnoreWhitespace = true,
                    Async = true
                });

                while (await xr.ReadAsync())
                {
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        if (xr.LocalName != "presetShapeDefinitons")
                        {
                            var item = await LoadPresetShapeDefinitionAsync(xr);
                            _shapeDefinitions.Add(item.Style, item);
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                throw (new IOException("Cannot preset shape definitions file:presetShapeDefinitions.xml", ex));
            }
        }

        public static async Task<bool> LoadPresetShapeDefinitionFromXmlAsyncTriangle()
        {
            _shapeDefinitions = new Dictionary<eShapeStyle, ShapeDefinition>();

            var xmlFile = Directory.GetCurrentDirectory() + "\\resource\\triangleOnly.xml";
            var fs = new FileStream(xmlFile, FileMode.Open, FileAccess.Read);
            var xr = XmlReader.Create(fs, new XmlReaderSettings()
            {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreWhitespace = true,
                Async = true
            });
            while (await xr.ReadAsync())
            {
                if (xr.NodeType == XmlNodeType.Element)
                {
                    if (xr.LocalName == "triangle" && xr.NodeType != XmlNodeType.EndElement)
                    {
                        var item = await LoadPresetShapeDefinitionAsync(xr);
                        _shapeDefinitions.Add(item.Style, item);
                    }
                }
            }
            return true;
        }

        private static async Task<ShapeDefinition> LoadPresetShapeDefinitionAsync(XmlReader xr)
        {
            if (Enum.TryParse<eShapeStyle>(xr.LocalName, true, out var style))
            {
                var psd = new ShapeDefinition()
                {
                    Style = style
                };
                await LoadFromXmlAsync(psd, xr);
                return psd;
            }
            throw new InvalidOperationException();
        }
        private static async Task LoadFromXmlAsync(ShapeDefinition psd, XmlReader xr)
        {
            while (await xr.ReadAsync())
            {
                if (xr.NodeType == XmlNodeType.Element)
                {
                    switch (xr.LocalName)
                    {
                        case "avLst":
                            psd.ShapeAdjustValues = await LoadShapeGuidesAsync(xr);
                            break;
                        case "gdLst":
                            psd.ShapeGuides = await LoadShapeGuidesAsync(xr);
                            break;
                        case "ahLst":
                            psd.ShapeAdjustHandles = await LoadAdjustHandleAsync(xr);
                            break;
                        case "cxnLst":
                            psd.ShapeConnectionSite = await LoadConnectionLstAsync(xr);
                            break;
                        case "rect":
                            psd.TextBoxRect = new TextBoxRect() { TopName = xr.GetAttribute("t"), BottomName = xr.GetAttribute("b"), LeftName = xr.GetAttribute("l"), RightName = xr.GetAttribute("r") };
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

        private static async Task<TextBoxRect> LoadRectAsync(XmlReader xr)
        {
            TextBoxRect rect;
            while (await xr.ReadAsync())
            {
                if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "rect")
                {
                    rect = new TextBoxRect()
                    {
                        TopName = xr.GetAttribute("t"),
                        BottomName = xr.GetAttribute("b"),
                        LeftName = xr.GetAttribute("l"),
                        RightName = xr.GetAttribute("r")
                    };
                    return rect;
                }
            } 
            return null;
        }

        private static async Task<List<ShapeConnectionSite>> LoadConnectionLstAsync(XmlReader xr)
        {
            var shapeConnectionSite = new List<ShapeConnectionSite>();

            while (await xr.ReadAsync() && (xr.NodeType != XmlNodeType.EndElement && xr.LocalName != "cxnLst"))
            {
                var newConnection = new ShapeConnectionSite();

                var attrStr = xr.GetAttribute("ang");

                newConnection.Angle = xr.GetAttribute("ang");
                await xr.ReadAsync();

                newConnection.PositionCoordinate = new ShapePositionCoordinate() { X = xr.GetAttribute("x"), Y = xr.GetAttribute("y") };

                shapeConnectionSite.Add(newConnection);
                await xr.ReadAsync();

                if (xr.LocalName == "cxnLst" && xr.NodeType == XmlNodeType.EndElement)
                {
                    break;
                }
            }

            return shapeConnectionSite;
        }

        private static async Task<List<ShapeAdjustHandleBase>> LoadAdjustHandleAsync(XmlReader xr)
        {
            var l = new List<ShapeAdjustHandleBase>();
            var name = xr.LocalName;
            while (await xr.ReadAsync())
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

        private static async Task<List<ShapeGuide>> LoadShapeGuidesAsync(XmlReader xr)
        {
            var l = new List<ShapeGuide>();
            var name = xr.LocalName;
            while (await xr.ReadAsync())
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
#endif
    }
}