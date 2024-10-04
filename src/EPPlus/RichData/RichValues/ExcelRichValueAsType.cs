using OfficeOpenXml.RichData.RichValues.Errors;
using OfficeOpenXml.RichData.RichValues.LocalImage;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues
{
    internal class ExcelRichValueAsType
    {
        public ExcelRichValueAsType(ExcelRichValue richValue)
        {
            _richValue = richValue;
        }

        private readonly ExcelRichValue _richValue;

        public T Type<T>()
            where T : ExcelRichValue
        {
            return _richValue as T;
        }

        public  ErrorFieldRichValue ErrorField
        {
            get{ return _richValue as ErrorFieldRichValue; }
        }

        public ErrorPropagatedRichValue ErrorPropagated
        {
            get { return _richValue as ErrorPropagatedRichValue; }
        }

        public ErrorSpillRichValue ErrorSpill
        {
            get { return _richValue as ErrorSpillRichValue; }
        }

        public ErrorWithSubTypeRichValue ErrorWithSubType
        {
            get { return _richValue as ErrorWithSubTypeRichValue; }
        }

        public LocalImageRichValue LocalImage
        {
            get { return _richValue as LocalImageRichValue; }
        }

        public LocalImageAltTextRichValue LocalImageAltText
        {
            get { return _richValue as LocalImageAltTextRichValue; }
        }
    }
}
