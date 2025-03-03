using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class NumberFormatXml : INumberFormat
    {
        ExcelNumberFormatXml NumberFormat;

        internal NumberFormatXml(ExcelNumberFormatXml format)
        {
            NumberFormat = format;
        }

        public string NumberFormatString => NumberFormat.Format;

        public int NumberFormatID => NumberFormat.NumFmtId;

        public ExcelFormatTranslator Translator => NumberFormat.FormatTranslator;

        public ExcelIndexedColor? ColorId => Translator.NumFtColor;

        public bool HasValue
        {
            get
            {
                return !string.IsNullOrEmpty(NumberFormat.Id);
            }
        }
    }
}
