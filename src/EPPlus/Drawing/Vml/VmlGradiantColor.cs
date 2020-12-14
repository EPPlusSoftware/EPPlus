using System;
using System.Drawing;
namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// A color used in a vml gradient list
    /// </summary>
    public class VmlGradiantColor
    {
        public VmlGradiantColor(double percent, Color color)
        {
            if (percent < 0 || percent > 100)
            {
                throw new ArgumentOutOfRangeException("Percent must be in the interval of 0 to 100");
            }
            if(color.IsEmpty)
            {
                throw new ArgumentNullException("Parameter 'color' can't be empty");
            }
            Percent = percent;
            Color = color;
        }
        /// <summary>
        /// Percent position to insert the color. Range from 0-100
        /// </summary>
        public double Percent 
        {
            get;
            set;
        }
        public Color Color
        {
            get;
            set;
        }
    }
}