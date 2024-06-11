using System;

namespace OfficeOpenXml.Export.ToCollection.Exceptions
{
    /// <summary>
    /// Data convertion exception
    /// </summary>
    public class EPPlusDataTypeConvertionException : Exception
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="innerException"></param>
        internal EPPlusDataTypeConvertionException(string msg, Exception innerException) : base(msg, innerException)
        {
            
        }
    }
}
