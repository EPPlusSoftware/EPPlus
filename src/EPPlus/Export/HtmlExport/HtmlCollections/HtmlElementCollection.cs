/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
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
