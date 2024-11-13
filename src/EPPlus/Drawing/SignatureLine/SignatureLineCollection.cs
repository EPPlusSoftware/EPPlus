using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing
{
    internal class SignatureLineCollection : IEnumerable<ExcelSignatureLineStamp>
    {
        List<ExcelSignatureLineStamp> _list = new();

        internal SignatureLineCollection()
        {
        }

        /// <summary>
        /// Index operator, returns by 0-based index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ExcelSignatureLineStamp this[int index]
        {
            get 
            {
                if (_list[index].SignatureLineType == eSignatureLineType.Stamp)
                {
                    return _list[index];
                }
                else
                {
                    return _list[index] as ExcelSignatureLine;
                }
            }
            set 
            {
                _list[index] = value; 
            }
        }

        internal void Add(ExcelSignatureLineStamp chart)
        {
            _list.Add(chart);
        }

        public IEnumerator<ExcelSignatureLineStamp> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
    }
}
