using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Style.Interfaces
{
    public interface IExcelNumberFormat
    {        
        /// <summary>
        /// The numberformat string
        /// </summary>
        public string Format { get; }
        /// <summary>
        /// Number format Id
        /// </summary>
        public int NumFmtId { get; }
        /// <summary>
        /// If this numberformat is built in
        /// </summary>
        public bool BuildIn { get; }
    }
}
