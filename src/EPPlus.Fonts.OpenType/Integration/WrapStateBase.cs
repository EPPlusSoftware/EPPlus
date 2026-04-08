using EPPlus.Fonts.OpenType.TextShaping.DataHolders;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    internal abstract class WrapStateBase
    {
        public int LineStart { get; set; }
        public int WordStart { get; set; }
        public double CurrentLineWidth { get; set; }
        public double CurrentWordWidth { get; set; }

        /// <summary>
        /// Data holders for the individual lines
        /// </summary>
        public List<TextLineSimple> Lines = new List<TextLineSimple>();

        public TextLineSimple CurrentTextLine = new TextLineSimple();

        internal int CurrentFragmentIdx = 0;

        public bool IsCompleteWordReady(CharacterType charType, int currentPosition)
        {
            return (charType == CharacterType.Space || charType == CharacterType.EndOfText)
                   && WordStart < currentPosition;
        }
    }
}
