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

namespace OfficeOpenXml.Drawing.Vml
{
    internal class ExcelVmlDrawingCommentCollection : ExcelVmlDrawingBaseCollection, IEnumerable<ExcelVmlDrawingComment>, IEnumerator<ExcelVmlDrawingComment>, IDisposable
    {
        internal CellStore<ExcelVmlDrawingComment> _drawings;
        internal ExcelVmlDrawingCommentCollection(ExcelPackage pck, ExcelWorksheet ws, Uri uri) :
            base(pck, ws, uri)
        {
            _drawings = new CellStore<ExcelVmlDrawingComment>();
            if (uri == null)
            {
                VmlDrawingXml.LoadXml(CreateVmlDrawings());
            }
            else
            {
                AddDrawingsFromXml(ws);
            }
        }
        ~ExcelVmlDrawingCommentCollection()
        {
            _drawings.Dispose();
            _drawings = null;
        }
        protected void AddDrawingsFromXml(ExcelWorksheet ws)
        {
            var nodes = VmlDrawingXml.SelectNodes("//v:shape", NameSpaceManager);
            //var list = new List<IRangeID>();
            foreach (XmlNode node in nodes)
            {
                var rowNode = node.SelectSingleNode("x:ClientData/x:Row", NameSpaceManager);
                var colNode = node.SelectSingleNode("x:ClientData/x:Column", NameSpaceManager);
                if (rowNode != null && colNode != null)
                {
                    var row = int.Parse(rowNode.InnerText) + 1;
                    var col = int.Parse(colNode.InnerText) + 1;
                    //list.Add(new ExcelVmlDrawingComment(node, ws.Cells[row, col], NameSpaceManager));
                    _drawings.SetValue(row, col, new ExcelVmlDrawingComment(node, ws.Cells[row, col], NameSpaceManager));
                }
                else
                {
                    //list.Add(new ExcelVmlDrawingComment(node, ws.Cells[1, 1], NameSpaceManager));
                    _drawings.SetValue(1, 1, new ExcelVmlDrawingComment(node, ws.Cells[1, 1], NameSpaceManager));
                }
            }
            //list.Sort(new Comparison<IRangeID>((r1, r2) => (r1.RangeID < r2.RangeID ? -1 : r1.RangeID > r2.RangeID ? 1 : 0)));  //Vml drawings are not sorted. Sort to avoid missmatches.
            //_drawings = new RangeCollection(list);
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

            vml += "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">";
            vml += "<v:stroke joinstyle=\"miter\" />";
            vml += "<v:path gradientshapeok=\"t\" o:connecttype=\"rect\" />";
            vml += "</v:shapetype>";
            vml += "</xml>";

            return vml;
        }
        internal ExcelVmlDrawingComment Add(ExcelRangeBase cell)
        {
            XmlNode node = AddDrawing(cell);
            var draw = new ExcelVmlDrawingComment(node, cell, NameSpaceManager);
            _drawings.SetValue(cell._fromRow, cell._fromCol, draw);
            return draw;
        }
        private XmlNode AddDrawing(ExcelRangeBase cell)
        {
            int row = cell.Start.Row, col = cell.Start.Column;
            var node = VmlDrawingXml.CreateElement("v", "shape", ExcelPackage.schemaMicrosoftVml);

            int r = cell._fromRow, c = cell._fromCol;
            var prev = _drawings.PrevCell(ref r, ref c);
            if (prev)
            {
                var prevDraw = _drawings.GetValue(r, c);
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
        int _nextID = 0;
        /// <summary>
        /// returns the next drawing id.
        /// </summary>
        /// <returns></returns>
        internal string GetNewId()
        {
            if (_nextID == 0)
            {
                foreach (ExcelVmlDrawingComment draw in this)
                {
                    if (draw.Id.Length > 3 && draw.Id.StartsWith("vml"))
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
        internal ExcelVmlDrawingBase this[int row, int column]
        {
            get
            {
                return _drawings.GetValue(row, column);
            }
        }
        internal bool ContainsKey(int row, int column)
        {
            return _drawings.Exists(row, column);
        }
        internal int Count
        {
            get
            {
                return _drawings.Count;
            }
        }
        #region "Enumerator"
        CellStoreEnumerator<ExcelVmlDrawingComment> _enum;
        public IEnumerator<ExcelVmlDrawingComment> GetEnumerator()
        {
            Reset();
            return this;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            Reset();
            return this;
        }

        /// <summary>
        /// The current range when enumerating
        /// </summary>
        public ExcelVmlDrawingComment Current
        {
            get
            {
                return _enum.Current;
            }
        }

        /// <summary>
        /// The current range when enumerating
        /// </summary>
        object IEnumerator.Current
        {
            get
            {
                return _enum.Current;
            }
        }

        public bool MoveNext()
        {
            return _enum.Next();
        }

        public void Reset()
        {
            if (_enum != null) _enum.Dispose();
             _enum = new CellStoreEnumerator<ExcelVmlDrawingComment>(_drawings, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
        }
        void IDisposable.Dispose()
        {
            _enum.Dispose();
            _enum = null;
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }
        #endregion
    } 
}
