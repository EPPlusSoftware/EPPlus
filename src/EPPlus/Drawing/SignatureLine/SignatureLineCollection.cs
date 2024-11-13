using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Collection of signaturelines
    /// </summary>
    public class SignatureLineCollection : IEnumerable<ExcelSignatureLineStamp>
    {
        Dictionary<Guid, ExcelSignatureLineStamp> _dict = new();
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
                return _list[index];
            }
            set
            {
                _list[index] = value;
            }
        }
        //public ExcelSignatureLineStamp this[int index]
        //{
        //    get 
        //    {
        //        if (_list[index].SignatureLineType == eSignatureLineType.Stamp)
        //        {
        //            return _list[index];
        //        }
        //        else
        //        {
        //            return _list[index] as ExcelSignatureLine;
        //        }
        //    }
        //    set 
        //    {
        //        _list[index] = value; 
        //    }
        //}

        /// <summary>
        /// Get all signaturelines of stamp type
        /// </summary>
        public List<ExcelSignatureLineStamp> GetSignatureLineStamps()
        {
            var retList = new List<ExcelSignatureLineStamp>();
            foreach (var sline in _list)
            {
                if(sline.SignatureLineType == eSignatureLineType.Stamp)
                {
                    retList.Add(sline);
                }
            }
            return retList;
        }

        /// <summary>
        /// Get all signaturelines of line type
        /// </summary>
        public List<ExcelSignatureLine> GetSignatureLines()
        {
            var retList = new List<ExcelSignatureLine>();
            foreach (var sline in _list)
            {
                if (sline.SignatureLineType == eSignatureLineType.SignatureLine)
                {
                    retList.Add(sline as ExcelSignatureLine);
                }
            }
            return retList;
        }
        /// <summary>
        /// Get a signatureline by its id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public ExcelSignatureLineStamp GetSignatureLineById(Guid id)
        {
            return _dict[id];
        }

        /// <summary>
        /// Add a signatureline to the collection.
        /// </summary>
        /// <param name="signatureLine"></param>
        public void Add(ExcelSignatureLineStamp signatureLine)
        {
            _list.Add(signatureLine);
            _dict.Add(signatureLine.SetupID, signatureLine);
        }

        /// <summary>
        /// Remove a signatureline from the collection.
        /// </summary>
        /// <param name="signatureLine"></param>
        public void Remove(ExcelSignatureLineStamp sline)
        {
            _dict.Remove(sline.SetupID);
            _list.Remove(sline);
        }

        public void Clear()
        {
            _dict.Clear();
            _list.Clear();
        }

        /// <summary>
        /// Get enumerator
        /// </summary>
        /// <returns></returns>
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
