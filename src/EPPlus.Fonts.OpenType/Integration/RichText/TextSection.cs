using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace EPPlus.Fonts.OpenType.Integration.RichText
{
    internal class TextSection
    {
        int _startIdx;
        int _endIdx;

        internal string FullText { get { return _getFullText(); } }
        internal string Text { get { return _getText(_startIdx, _endIdx); } }

        Func<string> _getFullText;
        Func<int,int,string> _getText;

        /// <summary>
        /// Run starting char index in fulltext
        /// </summary>
        internal int FullTextStart { get { return _startIdx; } }
        internal int Length { get { return _endIdx - _startIdx; } }

        /// <summary>
        /// A section of text between start and end index
        /// </summary>
        /// <param name="getFullText">A function that gets the fulltext</param>
        /// <param name="getText">
        /// A function that takes the substring between
        /// start and end index on the fulltext</param>
        internal TextSection(int startIdx, int endIndex, Func<string> getFullText, Func<int, int, string> getText)
        {
            _startIdx = startIdx;
            _endIdx = endIndex;
            _getFullText = getFullText;
            _getText = getText;
        }
    }
}
