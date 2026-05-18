using System.Collections;

namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    internal class RenderTextRunCollection : IEnumerable<RenderTextRun>
    {
        private List<RenderTextRun> _textRunItems;

        public void Add(RenderTextRun item)
        {
            _textRunItems.Add(item);
        }

        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _textRunItems.Count;
            }
        }
        IEnumerator<RenderTextRun> IEnumerable<RenderTextRun>.GetEnumerator()
        {
            return _textRunItems.GetEnumerator();
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _textRunItems.GetEnumerator();
        }
        /// <summary>
        /// Returns a textrun at position
        /// </summary>
        /// <param name="PositionID">The position of the chart. 0-base</param>
        /// <returns></returns>
        public RenderTextRun this[int PositionID]
        {
            get
            {
                return (_textRunItems[PositionID]);
            }
        }
    }
}
