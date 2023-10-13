using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class FontDeclarationRules
    {
        internal FontDeclarationRules(IFont f, ExcelFont nf, TranslatorContext context) 
        {
            _f = f;
            _nf = nf;
            _fontExclude = context.Exclude.Font;
            _theme = context.Theme;
        }

        private readonly IFont _f;
        private readonly ExcelFont _nf;
        private readonly eFontExclude _fontExclude;
        private readonly ExcelTheme _theme;
        
        public bool HasFamily 
        {   
            get 
            {
                return (string.IsNullOrEmpty(_f.Name) == false && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Name) && _f.Name.Equals(_nf.Name) == false);
            } 
        }

        public bool HasSize
        {
            get
            {
                return (_f.Size > 0 && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Size) && _f.Size != _nf.Size);
            }
        }
        public bool HasColor
        {
            get
            {
                return (_f.Color != null && _f.Color.Exists && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Color) && HtmlUtils.ColorUtils.AreColorEqual(_f.Color, _nf.Color) == false);
            }
        }
        public bool HasBold
        {
            get
            {
                return (_f.Bold && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Bold) && _nf.Bold != _f.Bold);
            }
        }
        public bool HasItalic
        {
            get
            {
                return (_f.Italic && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Italic) && _nf.Italic != _f.Italic);
            }
        }
        public bool HasStrike
        {
            get
            {
                return (_f.Strike && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Strike) && _nf.Strike != _f.Strike);
            }
        }

        public bool HasUnderline
        {
            get
            {
                return (_f.UnderLineType != ExcelUnderLineType.None && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Underline) && _f.UnderLineType != _nf.UnderLineType);
            }
        }


    }
}
