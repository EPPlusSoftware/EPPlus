using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class DatabarSvg
    {
        //0 fill or #Gradient1
        //1 Stop left
        //2 Stop right
        //3 Border color
        internal const string DataBar = "<svg version='1.1' xmlns='http://www.w3.org/2000/svg' preserveAspectRatio='none'>" +
            "<defs>" +
            "<linearGradient id='Gradient1'><stop class='stop1' offset='0%' /><stop class='stop2' offset='90%' /></linearGradient>" +
            "<style> #rect1 {0} .stop1 {1} .stop2 {2} </style></defs>" +
            "<rect id='rect1' width='100%' height='100%' stroke='{3}' stroke-width='2px'/></svg>";
        internal const string AxisStripes = "<svg version='1.1' viewBox='0 0 15 100' xmlns='http://www.w3.org/2000/svg'><g fill='#140904'><rect id='stripe' width='15px' height='75%'/></g></svg>";
        internal const string AxisStripesColor = "<svg version='1.1' viewBox='0 0 15 100' xmlns='http://www.w3.org/2000/svg'><g fill='{0}'><rect id='stripe' width='15px' height='75%'/></g></svg>";


        internal static string GetConvertedDatabarString(Color databarColor, bool isGradient, Color? borderColor = null)
        {
            string svg = GetUncovertedDatabar(databarColor, isGradient, borderColor);
            return Convert.ToBase64String(Encoding.ASCII.GetBytes(svg));
        }
        internal static string GetConvertedAxisStripes()
        {
            return Convert.ToBase64String(Encoding.ASCII.GetBytes(AxisStripes));
        }

        internal static string GetConvertedAxisStripesWithColor(Color axisColor)
        {
            var colorAxis = string.Format(AxisStripesColor, GetColorCode(axisColor));
            return Convert.ToBase64String(Encoding.ASCII.GetBytes(colorAxis));
        }

        internal static string GetUncovertedDatabar(Color databarColor, bool isGradient, Color? borderColor = null)
        {

            string fill = "{ fill: ";

            fill += isGradient == true ? "url(#Gradient1)" : GetColorCode(databarColor);

            fill += "; }";

            string stopLeft = "{ stop-color: ";

            stopLeft += GetColorCode(databarColor);

            stopLeft += "; }";

            string stopRight = "{ stop-color: white; }";

            string borderColorStr = borderColor == null ? GetColorCode(Color.FromArgb(0, 0, 0, 0)) : GetColorCode(borderColor.Value);

            //if (isPositive)
            //{
            //    stopRight = "#ffff";
            //    stopLeft = GetColorCode(databarColor);
            //}
            //else
            //{
            //    stopRight = GetColorCode(databarColor);
            //    stopLeft = "#ffff";
            //}

            var stringRet = string.Format(DataBar, fill, stopLeft, stopRight, borderColorStr);
            return stringRet;
        }

        static string GetColorCode(Color color)
        {
            return "#" + color.ToArgb().ToString("x8").Substring(2);
        }

    }
}
