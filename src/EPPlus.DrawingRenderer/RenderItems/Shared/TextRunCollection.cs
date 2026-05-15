using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.DrawingRenderer.RenderItems
{
    internal class TextRunCollection : IEnumerable<TextRunItem>
    {
        private List<TextRunItem> _textRunItems;

        public void Add(TextRunItem item)
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
        IEnumerator<TextRunItem> IEnumerable<TextRunItem>.GetEnumerator()
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
        public TextRunItem this[int PositionID]
        {
            get
            {
                return (_textRunItems[PositionID]);
            }
        }
    }
}
