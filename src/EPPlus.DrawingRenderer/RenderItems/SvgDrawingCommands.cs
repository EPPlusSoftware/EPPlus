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
//using OfficeOpenXml.Drawing;
//using OfficeOpenXml.Utils.TypeConversion;
//using System;
//using System.Reflection;
//using System.Textbox.Json;
//using System.Windows;
//using System.Xml;
//namespace EPPlusImageRenderer.RenderItems
//{
//    public static class SvgDrawingCommands
//    {
//        public static void LoadShapes()
//        {
//            Shapes.Clear();
//            var br = new BinaryReader(new FileStream("c:\\temp\\Shapedrawing.bin", FileMode.Open));
//            while (br.BaseStream.Position < br.BaseStream.Length)
//            {
//                var shapeStyle = (eShapeStyle)br.ReadByte();
//                var list = new List<SvgRenderItem>();
//                var adj = ShapeAdjustments.ContainsKey(shapeStyle) ? ShapeAdjustments[shapeStyle] : null;
//                Shapes.Add(shapeStyle, list);
//                var ix = 0;
//                do
//                {
//                    var recType = br.ReadByte();
//                    if (recType == 0) break;
//                    SvgAdjustmentPoint adjPoint = null;
//                    if (adj != null && adj.ContainsKey(ix & 0x1FF))
//                    {
//                        adjPoint = adj[ix];
//                    }

//                    switch (recType)
//                    {
//                        case 1: //rect
//                            list.Add(ReadRect(br));
//                            break;
//                        case 2: //path
//                            list.Add(ReadPath(br));
//                            break;
//                    }
//                    ix++;
//                }
//                while (br.BaseStream.Position < br.BaseStream.Length);
//            }
//        }

//        private static SvgRenderPathItem ReadPath(BinaryReader br)
//        {
//            var fill = br.ReadByte();
//            var stroke = br.ReadByte();
//            var commandCount = br.ReadByte();
//            var item = new SvgRenderPathItem();
//            for (var i = 0; i < commandCount; i++)
//            {
//                var ct = (PathCommandType)br.ReadByte();
//                var itemCount = br.ReadInt16();
//                var l = new List<double>();
//                for (int j = 0; j < itemCount; j++)
//                {
//                    l.Add(br.ReadSingle());
//                }
//                var cmd = new PathCommands(ct, item, l.ToArray());
//                //if (adj != null && (adj.Commands == null || adj.Commands.Any(x => x.Index == i)))
//                //{
//                //    cmd.AdjustmentPoint = adj;
//                //    if (adj.Commands != null)
//                //    {
//                //        cmd.CommandIndex = adj.Commands.FindIndex(x => x.Index == i);
//                //    }
//                //}
//                item.Commands.Add(cmd);
//            }

//            item.FillColorSource = (PathFillMode)fill;
//            item.BorderColorSource = (PathFillMode)stroke;

//            return item;
//        }

//        private static SvgRenderRectItem ReadRect(BinaryReader br)
//        {
//            var fill = br.ReadByte();
//            var stroke = br.ReadByte();
//            var x = br.ReadSingle();
//            var y = br.ReadSingle();
//            var width = br.ReadSingle();
//            var height = br.ReadSingle();

//            return new SvgRenderRectItem() { Left = x, Top = y, Width = width, Height = height, FillColorSource = (PathFillMode)fill, BorderColorSource = (PathFillMode)stroke };
//        }
//        public static string SerializeShapeAdjustments()
//        {
//            var s = JsonSerializer.Serialize(ShapeAdjustments);
//            return new StringReader(s).ReadToEnd();
//        }
//        public static void DeSerializeShapeAdjustments(string json)
//        {
//            ShapeAdjustments.Clear();
//            LoadShapes();
//            //MessageBox.Show("Json Loaded");
//        }

//        public static void LoadAdjustments()
//        {
//            var path = new FileInfo(Assembly.GetExecutingAssembly().FullName).DirectoryName;
//            ShapeAdjustments = new Dictionary<eShapeStyle, Dictionary<int, SvgAdjustmentPoint>>();
//            foreach (var f in Directory.GetFiles(path + "\\Adjustments\\", "*.xml"))
//            {
//                ReadAdjustmentXml(f, ShapeAdjustments);
//            }
//        }

//        private static void ReadAdjustmentXml(string f, Dictionary<eShapeStyle, Dictionary<int, SvgAdjustmentPoint>> shapeAdjustments)
//        {
//            var xml = new XmlDocument();
//            xml.Load(f);

//            var style = (eShapeStyle)Enum.Parse(typeof(eShapeStyle), xml.SelectSingleNode("adjust/@type").Value);
//            var adjs = new Dictionary<int, SvgAdjustmentPoint>();
//            foreach (XmlElement objNode in xml.SelectNodes("adjust/objects/object"))
//            {
//                var adj = new SvgAdjustmentPoint();
//                adj.ItemIndex = int.Parse(objNode.GetAttribute("ix"));
//                adj.AdjustmentType = (AdjustmentType)Enum.Parse(typeof(AdjustmentType), objNode.GetAttribute("adjustmentType"));
//                foreach (XmlElement cmdNode in objNode.SelectNodes("commands/command"))
//                {
//                    var ix = int.Parse(cmdNode.GetAttribute("ix"));
//                    var cmd = new SvgCommand(ix);
//                    if (adj.Commands == null) adj.Commands = new List<SvgCommand>();
//                    adj.Commands.Add(cmd);
//                    foreach (XmlElement ptNode in cmdNode.SelectNodes("pt"))
//                    {
//                        var ptIx = short.Parse(ptNode.GetAttribute("ix"));
//                        var coordinate = new SvgCoordinate(ptIx);
//                        if (string.IsNullOrEmpty(ptNode.GetAttribute("adjustPointName"))==false)
//                        {
//                            coordinate.PointName = ptNode.GetAttribute("adjustPointName");
//                        }
//                        else
//                        {
//                            coordinate.PointName = "";
//                        }

//                        if (ConvertUtil.TryParseIntString(ptNode.GetAttribute("origin"), out int r))
//                        {
//                            coordinate.Origin = (short)r;
//                        }
//                        else
//                        {
//                            coordinate.Origin = 0;
//                        }
//                        if (Enum.TryParse<AdjustmentPointType>(ptNode.GetAttribute("type"), true, out var result))
//                        {
//                            coordinate.BulletType = result;
//                        }
//                        cmd.Coordinates.Add(ptIx, coordinate);
//                    }
//                }
//                adjs.Add(adj.ItemIndex, adj);
//            }
//            shapeAdjustments.Add(style, adjs);
//        }

//        internal static Dictionary<eShapeStyle, List<SvgRenderItem>> Shapes { get; } = new Dictionary<eShapeStyle, List<SvgRenderItem>>();
//        internal static Dictionary<eShapeStyle, Dictionary<int, SvgAdjustmentPoint>> ShapeAdjustments { get; set; } = new Dictionary<eShapeStyle, Dictionary<int, SvgAdjustmentPoint>>();
//    }

//}