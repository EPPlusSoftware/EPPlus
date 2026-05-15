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
using EPPlus.Graphics.Geometry;

namespace EPPlus.Graphics
{
    //REname and move class? Inherit from Transform? what do?
    public class Rect
    {

        public Rect()
        {
        }

        public Rect(double width, double height) : this()
        {
            Left = 0;
            Top = 0;
            Right = width;
            Bottom = height;
        }
        public Rect(double left, double top, double right, double bottom) : this()
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        /// <summary>
        /// Y pos (min)
        /// </summary>
        public virtual double Top { get; set; }

        /// <summary>
        /// Y pos (max)
        /// </summary>
        public virtual double Bottom { get; set; }

        /// <summary>
        /// X pos (min)
        /// </summary>
        public virtual double Left { get; set; }

        /// <summary>
        /// X pos (max)
        /// </summary>
        public virtual double Right { get; set; }

        /// <summary>
        /// Get or Set Width via the properties above
        /// </summary>
        public virtual double Width
        {
            get
            {
                return Right - Left;
            }
            set
            {
                Right = Left + value;
            }
        }

        /// <summary>
        /// Gets or sets height via the properties above
        /// </summary>
        public virtual double Height
        {
            get
            {
                return Bottom - Top;
            }
            set
            {
                Bottom = Top + value;
            }
        }

        /// <summary>
        /// Local position X
        /// X-position from parent Transform position
        /// </summary>
        public double X;

        /// <summary>
        /// Local position Y
        /// Y-position from parent Transform position
        /// </summary>
        public double Y;
    }
}
