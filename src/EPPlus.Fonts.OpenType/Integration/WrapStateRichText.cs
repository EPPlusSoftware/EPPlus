using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    internal class WrapStateRichText : WrapStateBase
    {
        internal LineFragment LineFrag = null;
        internal int CharIdxRt = 0;
        internal int CharIdxWithinOriginal = 0;

        public WrapStateRichText(double lineWidth) 
        {
            CurrentLineWidth = lineWidth;
        }

        internal void EndCurrentTextLine()
        {
            if (CurrentTextLine.InternalLineFragments.Contains(LineFrag) == false)
            {
                CurrentTextLine.InternalLineFragments.Add(LineFrag);
            }
            Lines.Add(CurrentTextLine);
        }

        internal void EndCurrentTextLineAndIntializeNext(int startIdxOfNewFragment)
        {
            var nextLine = new TextLineSimple();

            if (_fragmentsForNextLine == null)
            {
                EndCurrentTextLine();
                var spcWidthTemp = LineFrag.SpaceWidth;
                LineFrag = new LineFragment(CurrentFragmentIdx, startIdxOfNewFragment, CharIdxRt, CharIdxWithinOriginal);
                LineFrag.SpaceWidth = spcWidthTemp;
            }
            else
            {
                //_fragmentsForNextLine.Add(LineFrag);
                nextLine.InternalLineFragments = _fragmentsForNextLine;

                //LineFrag.StartIdx = ;
                Lines.Add(CurrentTextLine);
                _fragmentsForNextLine = null;
            }
            CurrentTextLine = nextLine;
        }

        int _rtIdxAtWordStart = -1;
        int _listIdxWithinLine = -1;
        double _lineFragWidthAtWordStart = -1;

        internal void SetAndLogWordStartState(int wordStart)
        {
            WordStart = wordStart;
            CurrentWordWidth = 0;

            _rtIdxAtWordStart = CurrentFragmentIdx;
            _lineFragWidthAtWordStart = LineFrag.Width;

            if(CurrentTextLine.InternalLineFragments.Count == 0)
            {
                _listIdxWithinLine = 0;
                return;
            }

            _listIdxWithinLine = CurrentTextLine.InternalLineFragments.Count - 1;

            if (CurrentTextLine.InternalLineFragments[_listIdxWithinLine].FragmentIndex < _rtIdxAtWordStart)
            {
                //When the word begins we are on a fragment that has not yet been added to the list.
                //It will be the next index when added
                _listIdxWithinLine += 1;
            }
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
            if(_rtIdxAtWordStart == CurrentFragmentIdx)
            {
                //If we are On the fragment we have not added it yet
                //Do so before splitting
                CurrentTextLine.InternalLineFragments.Add(LineFrag);
            }

            var origFragment = CurrentTextLine.InternalLineFragments[_listIdxWithinLine];

            var lenAtBreak = CharIdxWithinOriginal - origFragment.StartOriginal;

            var resultingFragment = CurrentTextLine.SplitAndGetLeftoverLineFragment(ref origFragment, _lineFragWidthAtWordStart, CharIdxRt, CharIdxWithinOriginal);
            CurrentTextLine.InternalLineFragments[_listIdxWithinLine] = origFragment;

            _fragmentsForNextLine = new List<LineFragment>();

            //Iterate backwards from back of list until we hit fragment
            for (int i = CurrentTextLine.InternalLineFragments.Count()-1; i > _listIdxWithinLine; i--)
            {
                //The current fragment's startidx is affected by the split if it is not the first in new line
                if (CurrentTextLine.InternalLineFragments[i].StartIdx != 0)
                {
                    CurrentTextLine.InternalLineFragments[i].StartIdx -= WordStart + 1;
                }
                //Add fragment to the new list
                _fragmentsForNextLine.Insert(0, CurrentTextLine.InternalLineFragments[i]);
                //Remove it from the old
                CurrentTextLine.InternalLineFragments.RemoveAt(i);
            }

            if (_rtIdxAtWordStart != CurrentFragmentIdx)
            {
                if (resultingFragment.Width != 0)
                {
                    //We also insert the fragment we've split out
                    //Unless the leftover is an empty fragment (this happens when we wrap on a space that had to be trimmed)
                    _fragmentsForNextLine.Insert(0, resultingFragment);
                }
                //The current fragment's startidx is affected by the split if it is not the first in new line
                if (LineFrag.StartIdx != 0)
                {
                    LineFrag.StartIdx -= WordStart + 1;
                }
            }
            else
            {
                LineFrag = resultingFragment;
            }

            _rtIdxAtWordStart = -1;
            _listIdxWithinLine = -1;
            _lineFragWidthAtWordStart = -1;
        }
    }
}
