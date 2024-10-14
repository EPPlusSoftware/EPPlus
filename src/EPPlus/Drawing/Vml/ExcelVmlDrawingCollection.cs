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
using System.Xml;
using System.Collections;
using System.Globalization;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Drawing.Controls;
using System.Text;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Constants;
using System.IO;
using OfficeOpenXml.Drawing.OleObject;
using System.Data.SqlTypes;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.DigitalSignatures;

namespace OfficeOpenXml.Drawing.Vml
{
    internal class ExcelVmlDrawingCollection
        : ExcelVmlDrawingBaseCollection, IEnumerable<ExcelVmlDrawingBase>, IDisposable, IPictureRelationDocument
    {
        internal CellStore<int> _drawingsCellStore;
        internal Dictionary<string, int> _drawingsDict = new Dictionary<string, int>();
        internal List<ExcelVmlDrawingBase> _drawings = new List<ExcelVmlDrawingBase>();
        public List<ExcelSignatureLine> SignatureLines = new List<ExcelSignatureLine>();
        Dictionary<string, HashInfo> _hashes = new Dictionary<string, HashInfo>();

        internal ExcelVmlDrawingCollection(ExcelWorksheet ws, Uri uri) :
            base(ws, uri, "d:legacyDrawing/@r:id")
        {
            _drawingsCellStore = new CellStore<int>();
            if (uri == null)
            {
                VmlDrawingXml.LoadXml(CreateVmlDrawings());
            }
            else
            {
                AddDrawingsFromXml(ws);
            }
        }
        ~ExcelVmlDrawingCollection()
        {
            _drawingsCellStore?.Dispose();
            _drawingsCellStore = null;
        }
        protected internal void AddDrawingsFromXml(ExcelWorksheet ws)
        {
            var nodes = VmlDrawingXml.SelectNodes("//v:shape", NameSpaceManager);
            //var list = new List<IRangeID>();
            foreach (XmlNode node in nodes)
            {
                var objectType = node.SelectSingleNode("x:ClientData/@ObjectType", NameSpaceManager)?.Value;
                ExcelVmlDrawingBase vmlDrawing;
                switch (objectType)
                {
                    case "Drop":
                    case "List":
                    case "Button":
                    case "GBox":
                    case "Label":
                    case "Checkbox":
                    case "Spin":
                    case "Radio":
                    case "EditBox":
                    case "Dialog":
                    case "Scroll":
                        vmlDrawing = new ExcelVmlDrawingControl(_ws, node, NameSpaceManager);
                        _drawings.Add(vmlDrawing);
                        break;
                    case "Pict":
                        vmlDrawing = new ExcelVmlDrawingPicture(node, NameSpaceManager, ws);
                        _drawings.Add(vmlDrawing);
                        break;
                    default:    //Comments
                        var rowNode = node.SelectSingleNode("x:ClientData/x:Row", NameSpaceManager);
                        var colNode = node.SelectSingleNode("x:ClientData/x:Column", NameSpaceManager);
                        int row, col;
                        if (rowNode != null && colNode != null)
                        {
                            row = int.Parse(rowNode.InnerText) + 1;
                            col = int.Parse(colNode.InnerText) + 1;
                        }
                        else
                        {
                            row = 1;
                            col = 1;
                        }
                        vmlDrawing = new ExcelVmlDrawingComment(node, ws.Cells[row, col], NameSpaceManager);
                        _drawings.Add(vmlDrawing);
                        _drawingsCellStore.SetValue(row, col, _drawings.Count-1);
                        break;
                }
                var id = string.IsNullOrEmpty(vmlDrawing.SpId) ? vmlDrawing.Id : vmlDrawing.SpId;
                int x = 2;
                if(_drawingsDict.ContainsKey(id))
                {
                    while(_drawingsDict.ContainsKey($"{id}-{x}"))
                    {
                        x++;
                    }
                    id = $"{id}-{x}";
                    if (string.IsNullOrEmpty(vmlDrawing.SpId))
                    {
                        vmlDrawing.Id= id;
                    }
                    else
                    {
                        vmlDrawing.SpId = id;
                    }
                }
                _drawingsDict.Add(id, _drawings.Count - 1);
            }
        }

        private string CreateVmlDrawings()
        {
            string vml = string.Format("<xml xmlns:v=\"{0}\" xmlns:o=\"{1}\" xmlns:x=\"{2}\">",
                ExcelPackage.schemaMicrosoftVml,
                ExcelPackage.schemaMicrosoftOffice,
                ExcelPackage.schemaMicrosoftExcel);

            vml += "<o:shapelayout v:ext=\"edit\">";
            vml += "<o:idmap v:ext=\"edit\" data=\"1\"/>";
            vml += "</o:shapelayout>";
            vml += "<v:shapetype path=\"m,l,21600r21600,l21600,xe\" o:spt=\"201\" coordsize=\"21600,21600\" id=\"_x0000_t201\">"; 
            vml += "<v:stroke joinstyle=\"miter\"/>";
            vml += "<v:path o:connecttype=\"rect\" fillok=\"f\" strokeok=\"f\" o:extrusionok=\"f\" shadowok=\"f\"/>";
            vml += "<o:lock v:ext=\"edit\" shapetype=\"t\"/>";
            vml += "</v:shapetype>";
            vml += "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">";
            vml += "<v:stroke joinstyle=\"miter\" />";
            vml += "<v:path gradientshapeok=\"t\" o:connecttype=\"rect\" />";
            vml += "</v:shapetype>";
            vml += "</xml>";

            return vml;
        }
        internal ExcelVmlDrawingComment AddComment(ExcelRangeBase cell)
        {
            XmlNode node = AddCommentDrawing(cell);
            var draw = new ExcelVmlDrawingComment(node, cell, NameSpaceManager);
            _drawings.Add(draw);
            _drawingsCellStore.SetValue(cell._fromRow, cell._fromCol, _drawings.Count-1);
            return draw;
        }
        private XmlNode AddCommentDrawing(ExcelRangeBase cell)
        {
            CreateVmlPart(false); //Create the vml part to be able to create related parts (like blip fill images).
            int row = cell.Start.Row, col = cell.Start.Column;
            var node = VmlDrawingXml.CreateElement("v", "shape", ExcelPackage.schemaMicrosoftVml);
            int r = cell._fromRow, c = cell._fromCol;
            var prev = _drawingsCellStore.PrevCell(ref r, ref c);
            if (prev)
            {                
                var prevDraw = _drawings[_drawingsCellStore.GetValue(r, c)];
                prevDraw.TopNode.ParentNode.InsertBefore(node, prevDraw.TopNode);
            }
            else
            {
                VmlDrawingXml.DocumentElement.AppendChild(node);
            }

            node.SetAttribute("id", GetNewId());
            node.SetAttribute("type", "#_x0000_t202");
            node.SetAttribute("style", "position:absolute;z-index:1; visibility:hidden");
            //node.SetAttribute("style", "position:absolute; margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1; visibility:hidden"); 
            node.SetAttribute("fillcolor", "#ffffe1");
            node.SetAttribute("insetmode", ExcelPackage.schemaMicrosoftOffice, "auto");

            string vml = "<v:fill color2=\"#ffffe1\" />";
            vml += "<v:shadow on=\"t\" color=\"black\" obscured=\"t\" />";
            vml += "<v:path o:connecttype=\"none\" />";
            vml += "<v:textbox style=\"mso-direction-alt:auto\">";
            vml += "<div style=\"text-align:left\" />";
            vml += "</v:textbox>";
            vml += "<x:ClientData ObjectType=\"Note\">";
            vml += "<x:MoveWithCells />";
            vml += "<x:SizeWithCells />";
            vml += string.Format("<x:Anchor>{0}, 15, {1}, 2, {2}, 31, {3}, 1</x:Anchor>", col, row - 1, col + 2, row + 3);
            vml += "<x:AutoFill>False</x:AutoFill>";
            vml += string.Format("<x:Row>{0}</x:Row>", row - 1);
            vml += string.Format("<x:Column>{0}</x:Column>", col - 1);
            vml += "</x:ClientData>";

            node.InnerXml = vml;
            return node;
        }

        internal ExcelSignatureLine AddSignatureLine()
        {
            var xmlNode = VmlDrawingXml.ChildNodes[0];
            var shapeTypeNode = VmlDrawingXml.DocumentElement.SelectSingleNode("v:shapetype[@id='_x0000_t75']", NameSpaceManager);
            if (shapeTypeNode == null)
            {
                shapeTypeNode = VmlDrawingXml.CreateElement("v", "shapetype", ExcelPackage.schemaMicrosoftVml);
                var shapeTypeElement = ((XmlElement)shapeTypeNode);
                shapeTypeElement.SetAttribute("id", "_x0000_t75");
                shapeTypeElement.SetAttribute("coordsize", "21600,21600");
                shapeTypeElement.SetAttribute("spt", ExcelPackage.schemaMicrosoftOffice, "75");
                shapeTypeElement.SetAttribute("preferrelative", ExcelPackage.schemaMicrosoftOffice, "t");
                shapeTypeElement.SetAttribute("path", "m@4@5l@4@11@9@11@9@5xe");
                shapeTypeElement.SetAttribute("filled", "f");
                shapeTypeElement.SetAttribute("stroked", "f");

                var stroke = VmlDrawingXml.CreateElement("v", "stroke", ExcelPackage.schemaMicrosoftVml);
                stroke.SetAttribute("joinstyle", "miter");
                shapeTypeNode.AppendChild(stroke);

                ExcelSignatureLine.CreateFormulaElementAsChildOf(shapeTypeNode);

                var pathElement = VmlDrawingXml.CreateElement("v", "path", ExcelPackage.schemaMicrosoftVml);
                pathElement.SetAttribute("extrusionok", ExcelPackage.schemaMicrosoftOffice, "f");
                pathElement.SetAttribute("gradientshapeok", "t");
                pathElement.SetAttribute("connecttype", ExcelPackage.schemaMicrosoftOffice, "rect");

                shapeTypeElement.AppendChild(pathElement);

                var lockShapeTypeEl = VmlDrawingXml.CreateElement("o", "lock", ExcelPackage.schemaMicrosoftOffice);
                lockShapeTypeEl.SetAttribute("ext", ExcelPackage.schemaMicrosoftVml, "edit");
                lockShapeTypeEl.SetAttribute("aspectratio", "t");

                shapeTypeElement.AppendChild(lockShapeTypeEl);

                VmlDrawingXml.DocumentElement.AppendChild(shapeTypeElement);
            }

            Guid lineId = new Guid();
            XmlNode node = AddDigitalSignatureLineDrawing(lineId);
            VmlDrawingXml.DocumentElement.AppendChild(node);

            var sLine = new ExcelSignatureLine(_ws, node, NameSpaceManager, lineId);
            SignatureLines.Add(sLine);
            _drawings.Add(sLine);

            return sLine;
        }
        public XmlNode AddDigitalSignatureLineDrawing(Guid id)
        {
            CreateVmlPart(false); //Create the vml part to be able to create related parts (like signatureLine).
            //var vmlRel = Part.CreateRelationship(mediaUri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");

            var shapeElement = VmlDrawingXml.CreateElement("v", "shape", ExcelPackage.schemaMicrosoftVml);
            VmlDrawingXml.DocumentElement.AppendChild(shapeElement);

            shapeElement.SetAttribute("id", "_x0000_s1025");
            shapeElement.SetAttribute("type", "#_x0000_t75");
            shapeElement.SetAttribute("alt", "Microsoft Office Signature Line...");
            shapeElement.SetAttribute("style", "position:absolute;margin-left:0;margin-top:0;width:192pt;height:96pt;z-index:1");

            var imagedata = VmlDrawingXml.CreateElement("v", "imagedata", ExcelPackage.schemaMicrosoftVml);
            imagedata.SetAttribute("relid", ExcelPackage.schemaMicrosoftOffice, RelId ?? "rId1");
            imagedata.SetAttribute("title", ExcelPackage.schemaMicrosoftOffice, "");
            shapeElement.AppendChild(imagedata);

            var lockShapeEl = VmlDrawingXml.CreateElement("o", "lock", ExcelPackage.schemaMicrosoftOffice);
            lockShapeEl.SetAttribute("ext", ExcelPackage.schemaMicrosoftVml, "edit");
            lockShapeEl.SetAttribute("ungrouping", "t");
            lockShapeEl.SetAttribute("rotation", "t");
            lockShapeEl.SetAttribute("cropping", "t");
            lockShapeEl.SetAttribute("verticies", "t");
            lockShapeEl.SetAttribute("text", "t");
            lockShapeEl.SetAttribute("grouping", "t");
            shapeElement.AppendChild(lockShapeEl);

            var sigLine = VmlDrawingXml.CreateElement("o", "signatureline", ExcelPackage.schemaMicrosoftOffice);
            shapeElement.AppendChild(sigLine);

            sigLine.SetAttribute("ext", ExcelPackage.schemaMicrosoftVml, "edit");

            string provId = "00000000-0000-0000-0000-000000000000";

            sigLine.SetAttribute("id", $"{{{id}}}");
            sigLine.SetAttribute("provid", $"{{{provId}}}");
            sigLine.SetAttribute("suggestedsigner", ExcelPackage.schemaMicrosoftOffice, "");
            sigLine.SetAttribute("suggestedsigner2", ExcelPackage.schemaMicrosoftOffice, "");
            sigLine.SetAttribute("suggestedsigneremail", ExcelPackage.schemaMicrosoftOffice, "");
            sigLine.SetAttribute("issignatureline", "t");

            var clientData = VmlDrawingXml.CreateElement("x", "ClientData", ExcelPackage.schemaMicrosoftExcel);
            shapeElement.AppendChild(clientData);

            clientData.SetAttribute("ObjectType", "Pict");

            var sizeWCells = VmlDrawingXml.CreateElement("x", "SizeWithCells", ExcelPackage.schemaMicrosoftExcel);

            var anchor = VmlDrawingXml.CreateElement("x", "Anchor", ExcelPackage.schemaMicrosoftExcel);
            anchor.InnerText = "0, 0, 0, 0, 4, 0, 6, 8";

            var CF = VmlDrawingXml.CreateElement("x", "CF", ExcelPackage.schemaMicrosoftExcel);
            CF.InnerText = "Pict";

            var autoPict = VmlDrawingXml.CreateElement("x", "AutoPict", ExcelPackage.schemaMicrosoftExcel);

            clientData.AppendChild(sizeWCells);
            clientData.AppendChild(anchor);
            clientData.AppendChild(CF);
            clientData.AppendChild(autoPict);

            return shapeElement;
        }

        internal ExcelVmlDrawingPicture AddPicture(ExcelOleObject oleObject, string name, Uri mediaUri)
        {
            UpdateShapeTypeForOleObject();
            XmlNode node = AddOleObjectDrawing(oleObject, name, mediaUri);
            var draw = new ExcelVmlDrawingPicture(node, NameSpaceManager, _ws);
            _drawings.Add(draw);
            if (_drawingsDict.ContainsKey(draw.Id) == false)
            {
                _drawingsDict.Add(draw.Id, _drawings.Count - 1);
            }

            return draw;
        }

        private void UpdateShapeTypeForOleObject()
        {
            string vml = string.Format("<xml xmlns:v=\"{0}\" xmlns:o=\"{1}\" xmlns:x=\"{2}\">",
                ExcelPackage.schemaMicrosoftVml,
                ExcelPackage.schemaMicrosoftOffice,
                ExcelPackage.schemaMicrosoftExcel);

            vml += "<o:shapelayout v:ext=\"edit\">";
            vml += "<o:idmap v:ext=\"edit\" data=\"1\"/>";
            vml += "</o:shapelayout>";
            vml += "<v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\" o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">";
            vml += "<v:stroke joinstyle=\"miter\"/>";
            vml += "<v:formulas>";
            vml += "<v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>";
            vml += "<v:f eqn=\"sum @0 1 0\"/>";
            vml += "<v:f eqn=\"sum 0 0 @1\"/>";
            vml += "<v:f eqn=\"prod @2 1 2\"/>";
            vml += "<v:f eqn=\"prod @3 21600 pixelWidth\"/>";
            vml += "<v:f eqn=\"prod @3 21600 pixelHeight\"/>";
            vml += "<v:f eqn=\"sum @0 0 1\"/>";
            vml += "<v:f eqn=\"prod @6 1 2\"/>";
            vml += "<v:f eqn=\"prod @7 21600 pixelWidth\"/>";
            vml += "<v:f eqn=\"sum @8 21600 0\"/>";
            vml += "<v:f eqn=\"prod @7 21600 pixelHeight\"/>";
            vml += "<v:f eqn=\"sum @10 21600 0\"/>";
            vml += "</v:formulas>";
            vml += "<v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>";
            vml += "<o:lock v:ext=\"edit\" aspectratio=\"t\"/>";
            vml += "</v:shapetype>";
            vml += "</xml>";
            VmlDrawingXml.LoadXml(vml);
        }

        private XmlNode AddOleObjectDrawing(ExcelOleObject oleObject, string name, Uri mediaUri)
        {
            CreateVmlPart(false); //Create the vml part to be able to create related parts (like blip fill images).
            var shapeElement = VmlDrawingXml.CreateElement("v", "shape", ExcelPackage.schemaMicrosoftVml);
            VmlDrawingXml.DocumentElement.AppendChild(shapeElement);

            //Create relationship to image here
            var vmlRel = Part.CreateRelationship(mediaUri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");

            shapeElement.SetAttribute("id", "_x0000_s" + oleObject.Id);
            shapeElement.SetAttribute("type", "#_x0000_t75");
            shapeElement.SetAttribute("style", "position:absolute; margin-left:0;margin-top:0;width:61.5pt;height40.5pt:; z-index:1"); //SET VALUE BASED ON MEDIA ;mso-wrap-style:tight
            shapeElement.SetAttribute("filled", "t");
            shapeElement.SetAttribute("fillcolor", "window [65]");
            shapeElement.SetAttribute("stroked", "t");
            shapeElement.SetAttribute("strokecolor", "windowText [64]");
            shapeElement.SetAttribute("insetmode", ExcelPackage.schemaMicrosoftOffice, "auto");

            StringBuilder vml = new StringBuilder();
            vml.Append("<v:fill color2=\"window [65]\" />");
            vml.AppendFormat("<v:imagedata o:relid=\"{0}\" o:title=\"\" />", vmlRel.Id);
            vml.Append("<x:ClientData ObjectType=\"Pict\">");
            //vml.Append("<x:MoveWithCells />");
            vml.Append("<x:SizeWithCells />");
            vml.AppendFormat("<x:Anchor>0, 0, 0, 0, 1, 32, 3, 12</x:Anchor>"); //SET VALUE BASED ON MEDIA
            //vml.Append("<x:AutoFill>False</x:AutoFill>");
            vml.Append("<x:CF>Pict</x:CF>");
            vml.Append("<x:AutoPict/>");
            //vml.Append("<x:DDE />");
            //vml.Append("<x:Camera />");
            vml.Append("</x:ClientData>");

            shapeElement.InnerXml = vml.ToString();
            return shapeElement;
        }

        internal ExcelVmlDrawingControl AddControl(ExcelControl ctrl, string name)
        {
            XmlNode node = AddControlDrawing(ctrl, name);
            var draw = new ExcelVmlDrawingControl(_ws, node, NameSpaceManager);
            _drawings.Add(draw);
            if(_drawingsDict.ContainsKey(draw.Id) == false)
            {
                _drawingsDict.Add(draw.Id, _drawings.Count - 1);
            }
            return draw;
        }
        private XmlNode AddControlDrawing(ExcelControl ctrl, string name)
        {
            CreateVmlPart(false); //Create the vml part to be able to create related parts (like blip fill images).
            var shapeElement = VmlDrawingXml.CreateElement("v", "shape", ExcelPackage.schemaMicrosoftVml);

            VmlDrawingXml.DocumentElement.AppendChild(shapeElement);

            shapeElement.SetAttribute("spid", ExcelPackage.schemaMicrosoftOffice, "_x0000_s"+ctrl.Id);
            shapeElement.SetAttribute("id", name);
            //shapeElement.SetAttribute("id", $"{ctrl.ControlTypeString}_x{ctrl.Id}_1");
            shapeElement.SetAttribute("type", "#_x0000_t201");
            shapeElement.SetAttribute("style", "position:absolute;z-index:1;");
            shapeElement.SetAttribute("insetmode", ExcelPackage.schemaMicrosoftOffice, "auto");
            SetShapeAttributes(ctrl, shapeElement);

            var vml = new StringBuilder();
            vml.Append(GetVml(ctrl, shapeElement));
            vml.Append("<o:lock v:ext=\"edit\" rotation=\"t\"/>");
            vml.Append("<v:textbox style=\"mso-direction-alt:auto\" o:singleclick=\"f\">");
            if (ctrl is ExcelControlWithText textControl)
            {
                vml.Append($"<div style=\"text-align:center\"><font color=\"#000000\" size=\"{GetFontSize(ctrl)}\" face=\"{GetFontName(ctrl)}\"></font></div>");
            }
            vml.Append("</v:textbox>");
            vml.Append($"<x:ClientData ObjectType=\"{ctrl.ControlTypeString}\">");
            vml.Append(string.Format("<x:Anchor>{0}</x:Anchor>", ctrl.GetVmlAnchorValue()));
            vml.Append(GetVmlClientData(ctrl, shapeElement));
            vml.Append("<x:PrintObject>False</x:PrintObject>");
            vml.Append("<x:AutoFill>False</x:AutoFill>");
            if (ctrl.ControlType != eControlType.GroupBox)
            {
                vml.Append("<x:TextVAlign>Center</x:TextVAlign>");
            }

            vml.Append("</x:ClientData>");

            shapeElement.InnerXml = vml.ToString();
            return shapeElement;
        }
        private string GetFontName(ExcelControl ctrl)
        {
            if (ctrl.ControlType == eControlType.Button)
            {
                return "Calibri";
            }
            else
            {
                return "Segoe UI";
            }
        }

        private string GetFontSize(ExcelControl ctrl)
        {
            if (ctrl.ControlType == eControlType.Button)
            {
                return "220";
            }
            else
            {
                return "160";
            }
        }

        private string GetVmlClientData(ExcelControl ctrl, XmlElement shapeElement)
        {
            switch (ctrl.ControlType)
            {
                case eControlType.Button:
                    return "<x:TextHAlign>Center</x:TextHAlign>";
                case eControlType.CheckBox:
                case eControlType.GroupBox:
                    return "<x:SizeWithCells/><x:NoThreeD/>";
                case eControlType.RadioButton:
                    return "<x:SizeWithCells/><x:AutoLine>False</x:AutoLine><x:NoThreeD/><x:FirstButton/>";
                case eControlType.DropDown:
                    return "<x:SizeWithCells/><x:AutoLine>False</x:AutoLine><x:Val>0</x:Val><x:Min>0</x:Min><x:Max>0</x:Max><x:Inc>1</x:Inc><x:Page>1</x:Page><x:Dx>22</x:Dx><x:Sel>0</x:Sel><x:NoThreeD2/><x:SelType>Single</x:SelType><x:LCT>Normal</x:LCT><x:DropStyle>Combo</x:DropStyle>   <x:DropLines>8</x:DropLines>";
                case eControlType.ListBox:
                    return "<x:SizeWithCells/><x:AutoLine>False</x:AutoLine><x:Val>0</x:Val><x:Min>0</x:Min><x:Max>0</x:Max><x:Inc>1</x:Inc><x:Page>7</x:Page><x:Dx>22</x:Dx><x:Sel>0</x:Sel><x:NoThreeD2/><x:SelType>Single</x:SelType><x:LCT>Normal</x:LCT>";
                case eControlType.Label:
                    return "<x:AutoFill>False</x:AutoFill><x:AutoLine>False</x:AutoLine>";
                case eControlType.ScrollBar:
                    return "<x:SizeWithCells/><x:Val>0</x:Val><x:Min>0</x:Min><x:Max>100</x:Max><x:Inc>1</x:Inc><x:Page>10</x:Page><x:Dx>22</x:Dx>";
                case eControlType.SpinButton:
                    return "   <x:Val>0</x:Val><x:Min>0</x:Min><x:Max>30000</x:Max><x:Inc>1</x:Inc><x:Page>10</x:Page><x:Dx>22</x:Dx>";
                default:
                    return "";
            }
        }

        private string GetVml(ExcelControl ctrl, XmlElement shapeElement)
        {
            switch (ctrl.ControlType)
            {
                case eControlType.Button:
                    return "<v:fill o:detectmouseclick=\"t\" color2=\"buttonFace[67]\"/>";
                case eControlType.CheckBox:
                    return "<v:path fillok=\"t\" strokeok=\"t\" shadowok=\"t\"/>";
                default:
                    return "";
            }
        }

        private void SetShapeAttributes(ExcelControl ctrl, XmlElement shapeElement)
        {
            switch (ctrl.ControlType)
            {
                case eControlType.Button:
                    shapeElement.SetAttribute("fillcolor", "buttonFace [67]");
                    shapeElement.SetAttribute("strokecolor", "windowText [64]");
                    shapeElement.SetAttribute("button", ExcelPackage.schemaMicrosoftOffice, "t");
                    break;
                case eControlType.CheckBox:
                case eControlType.RadioButton:
                    shapeElement.SetAttribute("fillcolor", "windows [65]");
                    shapeElement.SetAttribute("strokecolor", "windowText [64]");
                    //shapeElement.SetAttribute("button", ExcelPackage.schemaMicrosoftOffice, "t");
                    shapeElement.SetAttribute("stroked", "f");
                    shapeElement.SetAttribute("filled", "f");
                    //style = "position:absolute; margin-left:15pt;margin-top:10.5pt;width:120.75pt;height:23.25pt;z-index:1; mso-wrap-style:tight" type = "#_x0000_t201" >
                    break;
                case eControlType.ListBox:
                case eControlType.DropDown:
                    shapeElement.SetAttribute("strokecolor", "windowText [64]");
                    shapeElement.SetAttribute("stroked", "f");
                    break;
                case eControlType.ScrollBar:
                    shapeElement.SetAttribute("strokecolor", "windowText [64]");
                    break;
                case eControlType.Label:
                    shapeElement.SetAttribute("fillcolor", "windows [65]");
                    shapeElement.SetAttribute("strokecolor", "windowText [64]");
                    shapeElement.SetAttribute("stroked", "f");
                    shapeElement.SetAttribute("filled", "f");
                    break;
            }

        }

        int _nextID = 0;
        /// <summary>
        /// returns the next drawing id.
        /// </summary>
        /// <returns></returns>
        internal string GetNewId()
        {
            if (_nextID == 0)
            {
                foreach (ExcelVmlDrawingBase draw in this)
                {
                    if (draw.Id.Length > 3 && draw.Id.StartsWith("vml", StringComparison.OrdinalIgnoreCase))
                    {
                        int id;
                        if (int.TryParse(draw.Id.Substring(3, draw.Id.Length - 3), System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out id))
                        {
                            if (id > _nextID)
                            {
                                _nextID = id;
                            }
                        }
                    }
                }
            }
            _nextID++;
            return "vml" + _nextID.ToString();
        }
        internal ExcelVmlDrawingBase this[string id]
        {
            get
            {
                if(_drawingsDict.ContainsKey(id))
                {
                    return _drawings[_drawingsDict[id]];
                }
                return null;
            }
        }

        internal ExcelVmlDrawingBase this[int row, int column]
        {
            get
            {
                return _drawings[_drawingsCellStore.GetValue(row, column)];
            }
        }
        internal bool ContainsKey(int row, int column)
        {
            return _drawingsCellStore.Exists(row, column);
        }
        internal int Count
        {
            get
            {
                return _drawings.Count;
            }
        }

        public ExcelPackage Package => _package;

        public Dictionary<string, HashInfo> Hashes => _hashes;

        public ZipPackagePart RelatedPart => Part;

        public Uri RelatedUri => Uri;
        #region "Enumerator"
        //CellStoreEnumerator<ExcelVmlDrawingComment> _enum;
        public IEnumerator<ExcelVmlDrawingBase> GetEnumerator()
        {
            //Reset();
            return _drawings.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            //Reset();
            return _drawings.GetEnumerator();
        }

        ///// <summary>
        ///// The current range when enumerating
        ///// </summary>
        //public ExcelVmlDrawingComment Current
        //{
        //    get
        //    {
        //        return _enum.Current;
        //    }
        //}

        ///// <summary>
        ///// The current range when enumerating
        ///// </summary>
        //object IEnumerator.Current
        //{
        //    get
        //    {
        //        return _enum.Current;
        //    }
        //}

        //public bool MoveNext()
        //{
        //    return _enum.Next();
        //}

        //public void Reset()
        //{
        //    if (_enum != null) _enum.Dispose();
        //     _enum = new CellStoreEnumerator<ExcelVmlDrawingComment>(_drawingsCellStore, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
        //}
        void IDisposable.Dispose()
        {
            _drawingsCellStore.Dispose();
        }

        //public void Dispose()
        //{
        //    throw new NotImplementedException();
        //}
        #endregion
    } 
}
