using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration.DataHolders
{
    public class LineFragmentCollection : List<LineFragment>
    {
        private string _text;

        public string FullText
        {
            get { return _text; }
            internal set { RichTextSubstrings.Clear(); _text = value; }
        }

        internal List<string> RichTextSubstrings { get; private set; }



        public LineFragmentCollection(string originalText)
        {
            FullText = originalText;
        }

        //Note: Array size never intended to be larger than 2
        List<int[]> StartEndPerFragment = new List<int[]>();

        
        /// <summary>
        /// Logs start and end idx per fragment
        /// Then adds fragment as regular list
        /// </summary>
        /// <param name="fragment"></param>
        public new void Add(LineFragment fragment)
        {
            var endIdx = FullText.Length - 1;
            if (Count != 0)
            {
                StartEndPerFragment.Last()[1] = fragment.StartIdx;
            }

            StartEndPerFragment.Add(new int[] { fragment.StartIdx, endIdx});
            base.Add(fragment);
        }

        public string GetLineFragmentText(LineFragment rtFragment)
        {
            if (this.Contains(rtFragment) == false)
            {
                throw new InvalidOperationException($"GetFragmentText failed. Cannot retrieve {rtFragment} since it is not part of this textLine: {this}");
            }

            if (string.IsNullOrEmpty(FullText))
            {
                return FullText;
            }

            var startIdx = rtFragment.StartIdx;

            var idxInLst = this.FindIndex(x => x == rtFragment);
            if (idxInLst == this.Count - 1)
            {
                return FullText.Substring(startIdx, FullText.Length - startIdx);
            }
            else
            {
                var endIdx = this[idxInLst + 1].StartIdx;
                return FullText.Substring(startIdx, endIdx - startIdx);
            }
        }

        internal void GenerateSubstrings()
        {
            RichTextSubstrings.Clear();

            //for (int i = 0; i < StartEndPerFragment.Count; i++)
            //{
            //    var startIdx = StartEndPerFragment[i][0];
            //    var endIdx = StartEndPerFragment[i][1];

            //    RichTextSubstrings.Add(FullText[1..5]);
            //}
            for (int i = 0; i < Count; i++)
            {
                RichTextSubstrings.Add(GetLineFragmentText(this[i]));
            }
        }
    }
}
