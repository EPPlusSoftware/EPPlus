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
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Defines the preset camera that is being used.
    /// </summary>
    public enum ePresetCameraType
    {
        /// <summary>
        /// No rotation
        /// </summary>
        None,
        /// <summary>
        /// Isometric Bottom Down
        /// </summary>
        IsometricBottomDown,
        /// <summary>
        /// Isometric Bottom Up
        /// </summary>
        IsometricBottomUp,
        /// <summary>
        /// Isometric Left Down
        /// </summary>
        IsometricLeftDown,
        /// <summary>
        /// Isometric Left Up
        /// </summary>
        IsometricLeftUp,
        /// <summary>
        /// Isometric Off Axis 1 Left
        /// </summary>
        IsometricOffAxis1Left,
        /// <summary>
        /// Isometric Off Axis 1 Right
        /// </summary>
        IsometricOffAxis1Right,
        /// <summary>
        /// Isometric Off Axis 1 Top
        /// </summary>
        IsometricOffAxis1Top,
        /// <summary>
        /// Isometric Off Axis 2 Left
        /// </summary>
        IsometricOffAxis2Left,
        /// <summary>
        /// Isometric Off Axis 2 Right
        /// </summary>
        IsometricOffAxis2Right,
        /// <summary>
        /// Isometric Off Axis 2 Top
        /// </summary>
        IsometricOffAxis2Top,
        /// <summary>
        /// Isometric Off Axis 3 Bottom
        /// </summary>
        IsometricOffAxis3Bottom,
        /// <summary>
        /// Isometric Off Axis 3 Left
        /// </summary>
        IsometricOffAxis3Left,
        /// <summary>
        /// Isometric Off Axis 3 Right
        /// </summary>
        IsometricOffAxis3Right,
        /// <summary>
        /// Isometric Off Axis 4 Bottom
        /// </summary>
        IsometricOffAxis4Bottom,
        /// <summary>
        /// Isometric Off Axis 4 Left
        /// </summary>
        IsometricOffAxis4Left,
        /// <summary>
        /// Isometric Off Axis 4 Right
        /// </summary>
        IsometricOffAxis4Right,
        /// <summary>
        /// Isometric Right Down
        /// </summary>
        IsometricRightDown,
        /// <summary>
        /// Isometric Right Up
        /// </summary>
        IsometricRightUp,
        /// <summary>
        /// Isometric Top Down
        /// </summary>
        IsometricTopDown,
        /// <summary>
        /// Isometric Top Up
        /// </summary>
        IsometricTopUp,
        /// <summary>
        /// Legacy Oblique Bottom
        /// </summary>
        LegacyObliqueBottom,
        /// <summary>
        /// Legacy Oblique Bottom Left
        /// </summary>
        LegacyObliqueBottomLeft,
        /// <summary>
        /// Legacy Oblique Bottom Right
        /// </summary>
        LegacyObliqueBottomRight,
        /// <summary>
        /// Legacy Oblique Front
        /// </summary>
        LegacyObliqueFront,
        /// <summary>
        /// 
        /// </summary>
        LegacyObliqueLeft,
        /// <summary>
        /// Legacy Oblique Right
        /// </summary>
        LegacyObliqueRight,
        /// <summary>
        /// Legacy Oblique Top
        /// </summary>
        LegacyObliqueTop,
        /// <summary>
        /// Legacy Oblique Top Left
        /// </summary>
        LegacyObliqueTopLeft,
        /// <summary>
        /// Legacy Oblique Top Right
        /// </summary>
        LegacyObliqueTopRight,
        /// <summary>
        /// Legacy Perspective Bottom
        /// </summary>
        LegacyPerspectiveBottom,
        /// <summary>
        /// Legacy Perspective Bottom Left
        /// </summary>
        LegacyPerspectiveBottomLeft,
        /// <summary>
        /// Legacy Perspective Bottom Right
        /// </summary>
        LegacyPerspectiveBottomRight,
        /// <summary>
        /// Legacy Perspective Front
        /// </summary>
        LegacyPerspectiveFront,
        /// <summary>
        /// Legacy Perspective Left
        /// </summary>
        LegacyPerspectiveLeft,
        /// <summary>
        /// Legacy Perspective Right
        /// </summary>
        LegacyPerspectiveRight,
        /// <summary>
        /// Legacy Perspective Top
        /// </summary>
        LegacyPerspectiveTop,
        /// <summary>
        /// Legacy Perspective Top Left
        /// </summary>
        LegacyPerspectiveTopLeft,
        /// <summary>
        /// Legacy Perspective Top Right
        /// </summary>
        LegacyPerspectiveTopRight,
        /// <summary>
        /// Oblique Bottom
        /// </summary>
        ObliqueBottom,
        /// <summary>
        /// Oblique Bottom Left
        /// </summary>
        ObliqueBottomLeft,
        /// <summary>
        /// Oblique Bottom Right
        /// </summary>
        ObliqueBottomRight,
        /// <summary>
        /// Oblique Left
        /// </summary>
        ObliqueLeft,
        /// <summary>
        /// Oblique Right
        /// </summary>
        ObliqueRight,
        /// <summary>
        /// Oblique Top
        /// </summary>
        ObliqueTop,
        /// <summary>
        /// Oblique Top Left
        /// </summary>
        ObliqueTopLeft,
        /// <summary>
        /// Oblique Top Right
        /// </summary>
        ObliqueTopRight,
        /// <summary>
        /// Orthographic Front
        /// </summary>
        OrthographicFront,
        /// <summary>
        /// Orthographic Above
        /// </summary>
        PerspectiveAbove,
        /// <summary>
        /// Perspective Above Left Facing
        /// </summary>
        PerspectiveAboveLeftFacing,
        /// <summary>
        /// Perspective Above Right Facing
        /// </summary>
        PerspectiveAboveRightFacing,
        /// <summary>
        /// Perspective Below
        /// </summary>
        PerspectiveBelow,
        /// <summary>
        /// Perspective Contrasting Left Facing
        /// </summary>
        PerspectiveContrastingLeftFacing,
        /// <summary>
        /// Perspective Contrasting Right Facing
        /// </summary>
        PerspectiveContrastingRightFacing,
        /// <summary>
        /// Perspective Front
        /// </summary>
        PerspectiveFront,
        /// <summary>
        /// Perspective Heroic Extreme Left Facing
        /// </summary>
        PerspectiveHeroicExtremeLeftFacing,
        /// <summary>
        /// Perspective Heroic Extreme Right Facing
        /// </summary>
        PerspectiveHeroicExtremeRightFacing,
        /// <summary>
        /// Perspective Heroic Left Facing
        /// </summary>
        PerspectiveHeroicLeftFacing,
        /// <summary>
        /// Perspective Right Facing
        /// </summary>
        PerspectiveHeroicRightFacing,
        /// <summary>
        /// Perspective Left
        /// </summary>
        PerspectiveLeft,
        /// <summary>
        /// Perspective Relaxed
        /// </summary>
        PerspectiveRelaxed,
        /// <summary>
        /// Perspective Relaxed Moderately
        /// </summary>
        PerspectiveRelaxedModerately,
        /// <summary>
        /// Perspective Right
        /// </summary>
        PerspectiveRight
    }
}