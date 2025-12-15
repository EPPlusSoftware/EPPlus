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
        internal Transform transform;

        internal Rect()
        {
            transform = new Transform();
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
        internal double Top
        {
            get { return transform.LocalPosition.Y; }
            set 
            {
                var tmpHeight = Height != 0 ? Height : 0;

                var currentPosition = transform.LocalPosition;
                currentPosition.Y = value;
                transform.LocalPosition = currentPosition;

                //Recalculate bottom position correctly
                if (tmpHeight != 0)
                {
                    Height = tmpHeight;
                }
            }
        }
        /// <summary>
        /// X pos (min)
        /// </summary>
        internal double Left
        {
            get { return transform.LocalPosition.X;}
            set 
            {
                var tmpWidth = Width != 0 ? Width : 0;

                var currentPosition = transform.LocalPosition;
                currentPosition.X = value;
                transform.LocalPosition = currentPosition;

                //Recalculate Right position correctly
                if(tmpWidth != 0)
                {
                    Width = tmpWidth;
                }
            }
        }

        /// <summary>
        /// X pos (max)
        /// </summary>
        internal double Right;
        /// <summary>
        /// Y pos (max)
        /// </summary>
        internal double Bottom;

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

        //Quick-access to underlying transform

        /// <summary>
        /// Local position X
        /// X-position from parent transform position
        /// </summary>
        internal double X { get { return Left; } set { Left = value; } }

        /// <summary>
        /// Local position Y
        /// Y-position from parent transform position
        /// </summary>
        internal double Y { get { return Top; } set { Top = value; } }

        //Gets global position x and y
        internal double GlobalX { get { return transform.Position.X; } }
        internal double GlobalY { get { return transform.Position.Y; } }
    }
}
