using EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    internal class WrapStateRichText : WrapStateBase
    {
        internal LineFragment LineFrag = null;

        public WrapStateRichText(double lineWidth) 
        {
            CurrentLineWidth = lineWidth;
        }

        internal void EndCurrentTextLine()
        {
            CurrentTextLine.LineFragments.Add(LineFrag);
            Lines.Add(CurrentTextLine);
        }

        internal void EndCurrentTextLineAndIntializeNext(int fragIdx, int startIdxOfNewFragment)
        {
            EndCurrentTextLine();

            CurrentTextLine = new TextLineSimple();

            if (_fragmentsForNextLine == null)
            {
                LineFrag = new LineFragment(fragIdx, startIdxOfNewFragment);
            }
            else
            {
                //Set pre-calculated fragments
                CurrentTextLine.LineFragments = _fragmentsForNextLine;
                _fragmentsForNextLine = null;
            }
        }

        int _rtIdxAtWordStart = -1;
        int _lstIdxWithinLine = -1;
        double _prevWordWidthAtWordStart = -1;
        double _lineWidthAtWordStart = -1;
        double _lineFragWidthAtWordStart = -1;

        internal void SetAndLogWordStartState(int wordStart)
        {
            WordStart = wordStart;

            _prevWordWidthAtWordStart = CurrentWordWidth;
            CurrentWordWidth = 0;

            _lineWidthAtWordStart = CurrentLineWidth;
            _rtIdxAtWordStart = CurrentFragmentIdx;
            _lineFragWidthAtWordStart = LineFrag.Width;
            //Count okay since the fragment we are on has not yet been added to currentline at this point
            //So its index Will Be this.
            _lstIdxWithinLine = CurrentTextLine.LineFragments.Count;
        }

        internal int GetFragIdxAtWordStart()
        {
            return _rtIdxAtWordStart;
        }

        internal double GetFragmentWidthAtWordStart()
        {
            return _lineFragWidthAtWordStart;
        }

        List<LineFragment> _fragmentsForNextLine = null;

        internal void AdjustLineFragmentsForNextLine()
        {
            var origFragment = CurrentTextLine.LineFragments[_lstIdxWithinLine];

            var resultingFragment = CurrentTextLine.SplitAndGetLeftoverLineFragment(ref origFragment, _lineFragWidthAtWordStart);
            CurrentTextLine.LineFragments[_lstIdxWithinLine] = origFragment;

            _fragmentsForNextLine = new List<LineFragment>();

            //Iterate backwards from back of list until we hit fragment
            for (int i = CurrentTextLine.LineFragments.Count; i > _lstIdxWithinLine; i--)
            {
                //Add fragment to the new list
                _fragmentsForNextLine.Insert(0, CurrentTextLine.LineFragments[i-1]);
                //Remove it from the old
                CurrentTextLine.LineFragments.RemoveAt(i-1);
            }

            //We also insert the fragment we've split out
            _fragmentsForNextLine.Insert(0, resultingFragment);
        }
    }
}
