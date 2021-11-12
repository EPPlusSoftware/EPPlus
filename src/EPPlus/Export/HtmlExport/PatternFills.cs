/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/07/2021         EPPlus Software AB       Added Html Export
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class PatternFills
    {
        internal const string Dott75 =                    "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='2'><rect width='4' height='2' fill='{1}'/><rect x='0' y='0' width='1' height='1' fill='{0}'/><rect x='2' y='1' width='1' height='1' fill='{0}'/></svg>";
        internal const string Dott50 =                    "<svg xmlns='http://www.w3.org/2000/svg' width='2' height='2'><rect width='2' height='2' fill='{0}'/><rect x='0' y='0' width='1' height='1' fill='{1}'/><rect x='1' y='1' width='1' height='1' fill='{1}'/></svg>";
        internal const string Dott25 =                    "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='2'><rect width='4' height='2' fill='{0}'/><rect x='2' y='0' width='1' height='1' fill='{1}'/><rect x='0' y='1' width='1' height='1' fill='{1}'/></svg>";
        internal const string Dott12_5 =                  "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='3' y='1' width='1' height='1' fill='{1}'/><rect x='1' y='3' width='1' height='1' fill='{1}'/></svg>";
        internal const string Dott6_25 =                  "<svg xmlns='http://www.w3.org/2000/svg' width='8' height='4'><rect width='8' height='4' fill='{0}'/><rect x='7' y='0' width='1' height='1' fill='{1}'/><rect x='3' y='2' width='1' height='1' fill='{1}' /></svg>";
        internal const string HorizontalStripe =          "<svg xmlns='http://www.w3.org/2000/svg' width='1' height='4'><rect width='1' height='4' fill='{0}'/><rect x='0' y='1' width='1' height='2' fill='{1}'/></svg>";
        internal const string VerticalStripe =            "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='1'><rect width='4' height='1' fill='{0}'/><rect x='1' y='0' width='2' height='2' fill='{1}'/></svg>";
        internal const string ThinHorizontalStripe =      "<svg xmlns='http://www.w3.org/2000/svg' width='1' height='4'><rect width='1' height='4' fill='{0}'/><rect x='0' y='1' width='1' height='1' fill='{1}'/></svg>";
        internal const string ThinVerticalStripe =        "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='1'><rect width='4' height='1' fill='{0}'/><rect x='1' y='0' width='2' height='1' fill='{1}'/></svg>";
                                                         
        internal const string ReverseDiagonalStripe =     "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='2' y='0' width='2' height='1' fill='{1}'/><rect x='0' y='1' width='1' height='1' fill='{1}'/><rect x='3' y='1' width='1' height='1' fill='{1}'/><rect x='0' y='2' width='2' height='1' fill='{1}'/><rect x='1' y='3' width='2' height='1' fill='{1}'/></svg>";
        internal const string DiagonalStripe =            "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='2' y='0' width='2' height='1' fill='{1}'/><rect x='1' y='1' width='2' height='1' fill='{1}'/><rect x='0' y='2' width='2' height='1' fill='{1}'/><rect x='0' y='3' width='1' height='1' fill='{1}'/><rect x='3' y='3' width='1' height='1' fill='{1}'/></svg>";

        internal const string ThinReverseDiagonalStripe = "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='2' y='0' width='1' height='1' fill='{1}'/><rect x='3' y='1' width='1' height='1' fill='{1}'/><rect x='0' y='2' width='1' height='1' fill='{1}'/><rect x='1' y='3' width='1' height='1' fill='{1}'/></svg>";
        internal const string ThinDiagonalStripe =        "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='2' y='0' width='1' height='1' fill='{1}'/><rect x='1' y='1' width='1' height='1' fill='{1}'/><rect x='0' y='2' width='1' height='1' fill='{1}'/><rect x='3' y='3' width='1' height='1' fill='{1}'/></svg>";
        
        internal const string DiagonalCrosshatch =        "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='2' y='0' width='2' height='1' fill='{1}'/><rect x='2' y='0' width='2' height='2' fill='{1}'/><rect x='0' y='2' width='2' height='2' fill='{1}'/></svg>";                
        internal const string ThickDiagonalCrosshatch =   "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='2' y='0' width='2' height='1' fill='{1}'/><rect x='0' y='1' width='4' height='1' fill='{1}'/><rect x='0' y='2' width='2' height='1' fill='{1}'/><rect x='0' y='3' width='4' height='1' fill='{1}'/></svg>";        
        internal const string ThinHorizontalCrosshatch =  "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='3' y='0' width='1' height='1' fill='{1}'/><rect x='0' y='1' width='4' height='1' fill='{1}'/><rect x='3' y='2' width='1' height='1' fill='{1}'/><rect x='3' y='3' width='1' height='1' fill='{1}'/></svg>";        
        internal const string ThinDiagonalCrosshatch =    "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='0' y='0' width='1' height='1' fill='{1}'/><rect x='2' y='0' width='1' height='1' fill='{1}'/><rect x='3' y='1' width='1' height='1' fill='{1}'/><rect x='0' y='2' width='1' height='1' fill='{1}'/><rect x='2' y='2' width='1' height='1' fill='{1}'/><rect x='1' y='3' width='1' height='1' fill='{1}'/></svg>";

    }
}
