using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;

#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    abstract internal class TranslatorBase
    {
        protected List<Declaration> declarations;

        internal TranslatorBase() 
        {
            declarations = new List<Declaration>();
        }

        internal abstract List<Declaration> GenerateDeclarationList(TranslatorContext context);

#if !NET35
        internal async Task<List<Declaration>> GenerateDeclarationListAsync(TranslatorContext context)
        {
            await Task.Run(() => GenerateDeclarationList(context));
            return declarations;
        }
#endif


        protected void AddDeclaration(string name, params string[] values) 
        {
            declarations.Add(new Declaration(name, values));
        }

        /// <summary>
        /// Gets hexcode color for html as a string 
        /// </summary>
        /// <param name="c"></param>
        /// <param name="theme"></param>
        /// <returns></returns>
        internal static string GetColor(ExcelColorXml c, ExcelTheme theme)
        {
            Color ret;
            if (!string.IsNullOrEmpty(c.Rgb))
            {
                if (int.TryParse(c.Rgb, NumberStyles.HexNumber, null, out int hex))
                {
                    ret = Color.FromArgb(hex);
                }
                else
                {
                    ret = Color.Empty;
                }
            }
            else if (c.Theme.HasValue)
            {
                ret = Utils.ColorConverter.GetThemeColor(theme, c.Theme.Value);
            }
            else if (c.Indexed >= 0)
            {
                ret = ExcelColor.GetIndexedColor(c.Indexed);
            }
            else
            {
                //Automatic, set to black.
                ret = Color.Black;
            }
            if (c.Tint != 0)
            {
                ret = Utils.ColorConverter.ApplyTint(ret, Convert.ToDouble(c.Tint));
            }
            return "#" + ret.ToArgb().ToString("x8").Substring(2);
        }

        internal static bool AreColorEqual(ExcelColorXml c1, ExcelColor c2)
        {
            if (c1.Tint != c2.Tint) return false;
            if (c1.Indexed >= 0)
            {
                return c1.Indexed == c2.Indexed;
            }
            else if (string.IsNullOrEmpty(c1.Rgb) == false)
            {
                return c1.Rgb == c2.Rgb;
            }
            else if (c1.Theme != null)
            {
                return c1.Theme == c2.Theme;
            }
            else
            {
                return c1.Auto == c2.Auto;
            }
        }
    }
}
