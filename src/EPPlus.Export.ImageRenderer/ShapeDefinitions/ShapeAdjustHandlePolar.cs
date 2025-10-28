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

namespace EPPlusImageRenderer.ShapeDefinitions
{
    public class ShapeAdjustHandlePolar : ShapeAdjustHandleBase
    {
        public override ShapeAdjustHandleType AhType => ShapeAdjustHandleType.Polar;
        public string AngleAdjustmentGuide { get; set; }
        public string RadialAdjustmentGuide { get; set; }
        public object MaximumAngleAdjustment { get; set; }
        public object MaximumRadialAdjustment { get; set; }
        public object MinimumAngleAdjustment { get; set; }
        public object MinimumRadialAdjustment { get; set; }
        internal override ShapeAdjustHandleBase Clone()
        {
            return new ShapeAdjustHandlePolar()
            {
                AngleAdjustmentGuide = AngleAdjustmentGuide,
                RadialAdjustmentGuide = RadialAdjustmentGuide,
                MaximumAngleAdjustment = MaximumAngleAdjustment,
                MinimumAngleAdjustment = MinimumAngleAdjustment,
                MaximumRadialAdjustment = MaximumRadialAdjustment,
                MinimumRadialAdjustment = MinimumRadialAdjustment,
            };
        }

    }
}
