using System;

namespace OfficeOpenXml.Export.ToCollection.Exceptions
{
    public class EPPlusDataTypeConvertionException : Exception
    {
        internal EPPlusDataTypeConvertionException(string msg, Exception innerException) : base(msg, innerException)
        {
            
        }
    }
}
