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

using EPPlus.Fonts.OpenType.Utils;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml;
using System;
using System.Globalization;
using System.Text;

namespace EPPlusImageRenderer
{
    internal class PathCommands 
    {
        public PathCommands(PathCommandType type, SvgRenderItem item, params double[] coordinates)
        {
            Type = type;
            RenderItem = item;
            Coordinates = coordinates;
        }
        public SvgRenderItem RenderItem{ get; set;}
        public PathCommandType Type { get; }
        public double[] Coordinates { get; set; }
        public SvgAdjustmentPoint AdjustmentPoint { get; set; }
        public int CommandIndex { get; set; }

        public void Render(/*double width, double height, */StringBuilder sb)
        {
            sb.Append(Type.AsCommandChar());
            for (int i = 0; i < Coordinates.Length; i++)
            {
                string s;
                if (Type==PathCommandType.Arc && ((i & 7)==2 || (i & 7) == 3 || (i & 7) == 4)) // Arc flags are not coordinates, but should be rendered as integers
                {
                    s = Coordinates[i].ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    s = Coordinates[i].PointToPixelString();
                }
                sb.AppendFormat("{0} ", s);
            }
            if (Coordinates.Length > 0)
            {
                sb.Remove(sb.Length - 1, 1);
            }
        }
        internal virtual bool InPoints(double x)
        {
            return true;
        }      
        internal PathCommands Clone()
        {
            return new PathCommands(Type, RenderItem)
            {
                Coordinates = (double[])Coordinates.Clone(),
            };
        }
    }

}