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

using EPPlusImageRenderer.RenderItems;
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

        public void Render(int width, int height, StringBuilder sb)
        {
            sb.Append(Type.AsCommandChar());
            for (int i = 0; i < Coordinates.Length; i++)
            {
                var x = Coordinates[i];
                if (Type == PathCommandType.Arc)
                {
                    switch(i)
                    {
                        case 0:                            
                        case 5:
                            x *= width;
                            break;
                        case 1:
                        case 6:
                            x *= height;
                            break;
                    }
                }
                else
                {
                    if (i % 2 == 0)
                    {
                        x *= width;
                    }
                    else
                    {
                        x *= height;
                    }
                    //}
                }
                sb.AppendFormat("{0} ", x.ToString(CultureInfo.InvariantCulture));
            }
            if (Coordinates.Length > 0)
            {
                sb.Remove(sb.Length - 1, 1);
            }
        }

        private double AdjustWithPoint(int width, int height, int i, double x)
        {
            var pt = AdjustmentPoint.Commands[CommandIndex].Coordinates[(short)i];
            var defPoint = 1;
            switch (pt.Type)
            {
                case AdjustmentPointType.Linear:
                    x = (x - (float)defPoint) + (float)defPoint;
                    break;
            }
            if (i % 2 == 0)
            {
                x *= width;
            }
            else
            {
                x *= height;
            }
            return x;
        }

        private double AdjustPointHalf(int width, int height, int i, double x, bool minus)
        {
            if (i % 2 == 0)
            {
                if (width > height)
                {
                    x *= Math.Max(width, height);
                    if (minus)
                    {
                        x -= (Math.Abs(width - height) / 2);
                    }
                    else
                    {
                        x += (Math.Abs(width - height) / 2);
                    }
                }
                else
                {
                    x *= width;
                }
            }
            else
            {
                if (height > width)
                {
                    x *= Math.Max(width, height);
                    if (minus)
                    {
                        x -= (Math.Abs(width - height) / 2);
                    }
                    else
                    {
                        x += (Math.Abs(width - height) / 2);
                    }
                }
                else
                {
                    x *= height;
                }
            }

            return x;
        }

        private static double AdjustToWidthHight(int width, int height, int i, double x)
        {
            if (i % 2 == 0)
            {
                if (width > height)
                {
                    x *= Math.Min(width, height);
                    x += (Math.Abs(width - height) / 2);
                }
                else
                {
                    x *= width;
                }
            }
            else
            {
                if (height > width)
                {
                    x *= Math.Min(width, height);
                    x += (Math.Abs(width - height) / 2);
                }
                else
                {
                    x *= height;
                }
            }

            return x;
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