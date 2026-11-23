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
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;

namespace OfficeOpenXml.Export.PdfExport.TextMapping
{
    internal static class PdfBorderReconciler
    {
        internal static void ReconcileSharedBorders(PdfCellCollection map)
        {
            for (int row = map.FromRow; row <= map.ToRow; row++)
            {
                for (int col = map.FromColumn; col <= map.ToColumn; col++)
                {
                    var cell = map[row, col];
                    if (cell == null || cell.Hidden || cell.Merged || cell.CellStyle == null) continue;
                    var cs = cell.CellStyle;
                    if (col < map.ToColumn)
                    {
                        var next = map[row, col + 1];
                        if (next != null && !next.Hidden && !next.Merged && next.CellStyle != null)
                        {
                            var ns = next.CellStyle;
                            if (!(IsDoubleEdge(cs.xfRight, cs.dxfRight) && IsDoubleEdge(ns.xfLeft, ns.dxfLeft)))
                            {
                                int here = EdgeRank(cs.xfRight, cs.dxfRight, cs.dxfRightElementOrder);
                                int there = EdgeRank(ns.xfLeft, ns.dxfLeft, ns.dxfLeftElementOrder);
                                if (here >= there) ns.SuppressLeft = true;
                                else cs.SuppressRight = true;
                            }
                        }
                    }
                    if (row < map.ToRow)
                    {
                        var below = map[row + 1, col];
                        if (below != null && !below.Hidden && !below.Merged && below.CellStyle != null)
                        {
                            var bs = below.CellStyle;
                            if (!(IsDoubleEdge(cs.xfBottom, cs.dxfBottom) && IsDoubleEdge(bs.xfTop, bs.dxfTop)))
                            {
                                int here = EdgeRank(cs.xfBottom, cs.dxfBottom, cs.dxfBottomElementOrder);
                                int there = EdgeRank(bs.xfTop, bs.dxfTop, bs.dxfTopElementOrder);
                                if (here >= there) bs.SuppressTop = true;
                                else cs.SuppressBottom = true;
                            }
                        }
                    }
                }
            }
        }

        private static bool IsDoubleEdge(ExcelBorderItem xf, ExcelDxfBorderItem dxf)
        {
            if (xf != null && xf.Style != ExcelBorderStyle.None) return xf.Style == ExcelBorderStyle.Double;
            if (dxf != null && dxf.Style.HasValue) return dxf.Style.Value == ExcelBorderStyle.Double;
            return false;
        }

        private static int EdgeRank(ExcelBorderItem xf, ExcelDxfBorderItem dxf, int elementOrder)
        {
            if (xf != null && xf.Style != ExcelBorderStyle.None)
                return PdfBorderStyleResolver.TableEdgeOrder.UserSet;
            if (dxf != null && dxf.Style.HasValue && dxf.Style.Value != ExcelBorderStyle.None)
                return elementOrder;
            return 0;
        }
    }
}
