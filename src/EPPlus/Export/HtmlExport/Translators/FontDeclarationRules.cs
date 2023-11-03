using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class FontDeclarationRules
    {
        bool _nfName = true;
        bool _nfSize = true;
        bool _nfColor = true;
        bool _nfBold = true;
        bool _nfItalic = true;
        bool _nfStrike = true;
        bool _nfUnderline = true;


        internal FontDeclarationRules(IFont f, ExcelFont nf, TranslatorContext context) 
        {
            _f = f;
            _nf = nf;
            _fontExclude = context.Exclude.Font;
            if (nf != null) 
            {
                _nfName = _f.Name.Equals(_nf.Name) == false;
                _nfSize = _f.Size != _nf.Size;
                _nfColor = f.Color.AreColorEqual(new StyleColorNormal(_nf.Color)) == false;
                _nfBold = _nf.Bold != _f.Bold;
                _nfItalic = _nf.Italic != _f.Italic;
                _nfStrike = _nf.Strike != _f.Strike;
                _nfUnderline = _f.UnderLineType != _nf.UnderLineType;
            }
        }

        private readonly IFont _f;
        private readonly ExcelFont _nf;
        private readonly eFontExclude _fontExclude;
        
        public bool HasFamily 
        {   
            get 
            {
                return (string.IsNullOrEmpty(_f.Name) == false && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Name) && _nfName);
            } 
        }

        public bool HasSize
        {
            get
            {
                return (_f.Size > 0 && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Size) && _nfSize);
            }
        }
        public bool HasColor
        {
            get
            {
                return (_f.Color != null && _f.Color.Exists && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Color) && _nfColor);
            }
        }
        public bool HasBold
        {
            get
            {
                return (_f.Bold && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Bold) && _nfBold);
            }
        }
        public bool HasItalic
        {
            get
            {
                return (_f.Italic && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Italic) && _nfItalic);
            }
        }
        public bool HasStrike
        {
            get
            {
                return (_f.Strike && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Strike) && _nfStrike);
            }
        }

        public bool HasUnderline
        {
            get
            {
                return (_f.UnderLineType != ExcelUnderLineType.None && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Underline) && _nfUnderline);
            }
        }
    }
}
