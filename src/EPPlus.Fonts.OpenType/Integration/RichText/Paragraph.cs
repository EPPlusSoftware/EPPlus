using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration.RichText
{
    /// <summary>
    /// A collection of StyleRuns
    /// </summary>
    internal class Paragraph : TextSection
    {
        private List<StyleRun> _styleRuns = new List<StyleRun>();
        internal Paragraph(int startIdx, int endIndex, Func<string> getFullText, Func<int, int, string> getText) : base(startIdx, endIndex, getFullText, getText)
        {
        }

        internal void AddStyleRun(StyleRun styleRun)
        {
            _styleRuns.Add(styleRun);
        }

        //Return array so original list cannot be changed
        internal StyleRun[] GetStyleRuns()
        {
            return _styleRuns.ToArray();
        }
    }
}
