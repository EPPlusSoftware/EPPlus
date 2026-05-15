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
using System.Xml;

namespace EPPlus.DrawingRenderer.ShapeDefinitions
{

    public abstract class ShapeAdjustHandleBase
    {
        public abstract ShapeAdjustHandleType AhType { get; }
        public ShapePositionCoordinate PositionCoordinate { get; set; }
        public static ShapeAdjustHandleXY CreateXy(XmlReader xr)
        {
            return new ShapeAdjustHandleXY()
            {
                HorizontalAdjustmentGuide = xr.GetAttribute("gdRefX"),
                VerticalAdjustmentGuide = xr.GetAttribute("gdRefY"),
                MinimumHorizontalAdjustment = xr.GetAttribute("minX"),
                MaximumHorizontalAdjustment = xr.GetAttribute("minY"),
                MinimumVerticalAdjustment = xr.GetAttribute("maxX"),
                MaximumVerticalAdjustment = xr.GetAttribute("maxY"),
            };
        }
        public static ShapeAdjustHandlePolar CreatePolar(XmlReader xr)
        {
            return new ShapeAdjustHandlePolar()
            {
                AngleAdjustmentGuide = xr.GetAttribute("gdRefAng"),
                RadialAdjustmentGuide = xr.GetAttribute("gdRefR"),
                MinimumAngleAdjustment = xr.GetAttribute("minAng"),
                MinimumRadialAdjustment = xr.GetAttribute("minR"),
                MaximumAngleAdjustment = xr.GetAttribute("maxAng"),
                MaximumRadialAdjustment = xr.GetAttribute("maxR"),
            };
        }

        public abstract ShapeAdjustHandleBase Clone();
    }
}
