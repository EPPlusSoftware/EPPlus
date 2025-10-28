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
    public class ShapeAdjustHandleXY : ShapeAdjustHandleBase
    {
        public override ShapeAdjustHandleType AhType => ShapeAdjustHandleType.XY;

        /// <summary>
        ///Specifies the name of the guide that is updated with the adjustment x position from this adjust handle.
        ///The possible values for this attribute are defined by the ST_GeomGuideName simple type (§20.1.10.28).
        /// </summary>
        public string HorizontalAdjustmentGuide { get; set; }
        /// <summary>
        /// Specifies the name of the guide that is updated with the adjustment y position from this adjust handle.
        /// </summary>
        public string VerticalAdjustmentGuide { get; set; }
        /// <summary>
        /// Specifies the minimum horizontal position that is allowed for this adjustment handle. 
        /// If this attribute is omitted, then it is assumed that this adjust handle cannot move in the x direction.
        /// That is the maxX and minX are equal.
        /// </summary>
        public object MinimumHorizontalAdjustment{ get; set; }
        /// <summary>
        /// Specifies the maximum horizontal position that is allowed for this adjustment handle. 
        /// If this attribute is omitted, then it is assumed that this adjust handle cannot move in the x direction.
        /// That is the maxX and minX are equal.
        /// </summary>
        public object MaximumHorizontalAdjustment { get; set; }
        /// <summary>
        /// Specifies the minimum vertical position that is allowed for this adjustment handle. 
        /// If this attribute is omitted, then it is assumed that this adjust handle cannot move in the y direction.That is the maxY and minY are equal.
        /// </summary>
        public object MinimumVerticalAdjustment { get; set; }
        /// <summary>
        /// Specifies the minimum vertical position that is allowed for this adjustment handle. 
        /// If this attribute is omitted, then it is assumed that this adjust handle cannot move in the y direction.
        /// That is the maxY and minY are equal.
        /// </summary>
        public object MaximumVerticalAdjustment { get; set; }
        internal override ShapeAdjustHandleBase Clone()
        {
            return new ShapeAdjustHandleXY()
            {
                HorizontalAdjustmentGuide = HorizontalAdjustmentGuide,
                VerticalAdjustmentGuide = VerticalAdjustmentGuide,
                MinimumHorizontalAdjustment = MinimumHorizontalAdjustment,
                MaximumHorizontalAdjustment = MaximumHorizontalAdjustment,
                MinimumVerticalAdjustment = MinimumVerticalAdjustment,
                MaximumVerticalAdjustment = MaximumVerticalAdjustment,
            };
        }
    }
}
