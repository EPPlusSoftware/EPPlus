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
using EPPlus.Graphics;
using System.Security.Cryptography;

namespace EPPlus.DrawingRenderer.RenderItems
{
    /// <summary>
    /// Describes how to position two rectangles relative to each other
    /// </summary>
    public enum RectangleAlignment
    {
        /// <summary>
        /// Bottom
        /// </summary>
        Bottom,
        /// <summary>
        /// Bottom Left
        /// </summary>
        BottomLeft,
        /// <summary>
        /// Bottom Right
        /// </summary>
        BottomRight,
        /// <summary>
        /// Center
        /// </summary>
        Center,
        /// <summary>
        /// Left
        /// </summary>
        Left,
        /// <summary>
        /// Right
        /// </summary>
        Right,
        /// <summary>
        /// Top
        /// </summary>
        Top,
        /// <summary>
        /// TopLeft
        /// </summary>
        TopLeft,
        /// <summary>
        /// TopRight
        /// </summary>
        TopRight
    }
    public enum TileFlipMode
    {
        /// <summary>
        /// Tiles are not flipped
        /// </summary>
        None,
        /// <summary>
        /// Tiles are flipped horizontally.
        /// </summary>
        X,
        /// <summary>
        /// Tiles are flipped horizontally and Vertically
        /// </summary>
        XY,
        /// <summary>
        /// Tiles are flipped vertically.
        /// </summary>
        Y
    }
    public class FillTile : RenderStyle
    {
        /// <summary>
        /// The direction(s) in which to flip the image.
        /// </summary>
        public TileFlipMode? FlipMode { get; set; }
        /// <summary>
        /// Where to align the first tile with respect to the shape.
        /// </summary>
        public RectangleAlignment? Alignment { get; set; }
        /// <summary>
        /// The ratio for horizontally scale
        /// </summary>
        public double HorizontalRatio { get; set; }
        /// <summary>
        /// The ratio for vertically scale
        /// </summary>
        public double VerticalRatio { get; set; }
        /// <summary>
        /// The horizontal offset after alignment
        /// </summary>
        public double HorizontalOffset { get; set; }
        /// <summary>
        /// The vertical offset after alignment
        /// </summary>
        public double VerticalOffset { get; set; }

        public override string GetKey()
        {
            return $"{FlipMode} {Alignment} {HorizontalRatio} {VerticalRatio} {HorizontalOffset} {VerticalOffset}";
        }
    }
    public class RenderBlipFill : RenderStyle
    {
        internal static SHA256 sha = SHA256.Create();
        public BoundingBox ImageBounds { get; set; } = new BoundingBox();
        public string ContentType { get; set; }
        public byte[] ImageBytes { get; set; }
        /// <summary>
        /// The image should be stretched to fill the target.
        /// </summary>
        public bool Stretch { get; set; } = false;
        public OffsetRectangle StretchOffset{ get; set; }
        public FillTile Tile{ get;set;}

        public override string GetKey()
        {
            var imageHash = Convert.ToBase64String(sha.ComputeHash(ImageBytes));
            return $"{ContentType} {imageHash} {Stretch} {StretchOffset.GetKey()} {Tile.GetKey()}";
        }
    }
}