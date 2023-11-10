using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.HtmlExport.HtmlCollections
{
    internal class HtmlElementCollection : IEnumerable<HTMLElement>
    {
        List<HTMLElement> _elements;

        internal List<HTMLElement> Elements => _elements;

        internal HtmlElementCollection()
        {
            _elements = new List<HTMLElement>();
        }

        IEnumerator<HTMLElement> IEnumerable<HTMLElement>.GetEnumerator()
        {
            for (int i = 0; i < _elements.Count; i++)
            {
                yield return _elements[i];
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _elements.GetEnumerator();
        }
    }
}
