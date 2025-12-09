using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.Nodes
{
    internal static class SvgElements
    {
        internal static readonly HashSet<string> VoidElements
        = new HashSet<string>
        {
            //{ "col" },
            //{ Img },
            //{ "input" }
        };

        internal static readonly HashSet<string> NoIndentElements
        = new HashSet<string>
        {
            //{ TableData },
            //{ TFoot },
            //{ TableHeader },
            //{ A },
            //{ Img }
        };

        //public const string Body = "body";
        //public const string Table = "table";
        //public const string Thead = "thead";
        //public const string TFoot = "tfoot";
        //public const string Tbody = "tbody";
        //public const string TableRow = "tr";
        //public const string TableHeader = "th";
        //public const string TableData = "td";
        //public const string A = "a";
        //public const string Span = "span";
        //public const string ColGroup = "colgroup";
        //public const string Img = "img";
    }
}
