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
using EPPlus.Graphics.Math;

namespace EPPlus.Graphics
{
    //REname and move class? Inherit from transform? what do?
    internal class Rect
    {

        internal Rect()
        {
        }

        internal Rect(double width, double height) : this()
        {
            Left = 0;
            Top = 0;
            Right = width;
            Bottom = height;
        }
        internal Rect(double left, double top, double right, double bottom) : this()
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        /// <summary>
        /// Y pos (min)
        /// </summary>
        internal double Top;

        /// <summary>
        /// Y pos (max)
        /// </summary>
        internal double Bottom;

        /// <summary>
        /// X pos (min)
        /// </summary>
        internal double Left;

        /// <summary>
        /// X pos (max)
        /// </summary>
        internal double Right;

        /// <summary>
        /// Get or Set Width via the properties above
        /// </summary>
        internal double Width
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
        internal double Height
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
        /// X-position from parent transform position
        /// </summary>
        internal double X;

        /// <summary>
        /// Local position Y
        /// Y-position from parent transform position
        /// </summary>
        internal double Y;
    }
}
