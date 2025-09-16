/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System.Xml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.EnumUtils;
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// A richtext part
    /// </summary>
   // public class ExcelTextRun : XmlHelper
   // {
   //     string _path;
   //     internal ExcelTextRun(XmlNamespaceManager ns, XmlNode topNode, string path) :
   //         base(ns, topNode)
   //     {
   //         _path = path;
   //         SchemaNodeOrder = new string[] { "ln", "noFill", "solidFill", "gradFill", "pattFill", "blipFill", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst", "highlight", "kumimoji", "lang", "altLang", "sz", "b", "i", "u", "strike", "kern", "cap", "spc", "normalizeH", "baseline", "noProof", "dirty", "err", "smtClean", "smtId", "bmk" };
   //         Properties = new ExcelRegularTextRun();
   //     }

   //     ExcelParagraphTextRun Properties;

   //     #region Attributes
   //     /// <summary>
   //     /// FontBold text
   //     /// </summary>
   //     public bool FontBold { get => Properties.Attributes.FontBold; set => Properties.Attributes.FontBold = value; }


   //     /// <summary>
   //     /// The baseline for both the superscript and subscript fonts in percentage
   //     /// </summary>
   //     public double Baseline { get => Properties.Attributes.Baseline; set => Properties.Attributes.Baseline = value; }

   //     /// <summary>
   //     /// The capitalization that is to be applied
   //     /// </summary>
   //     public eTextCapsType Capitalization { get => Properties.Attributes.Capitalization; set => Properties.Attributes.Capitalization = value; }
   //     //TODO: Dirty

   //     //TODO: err (spelling error)

   //     /// <summary>
   //     /// FontItalic text
   //     /// </summary>
   //     public bool FontItalic { get => Properties.Attributes.FontItalic; set => Properties.Attributes.FontItalic = value; }

   //     /// <summary>
   //     /// The minimum font size at which character kerning occurs
   //     /// </summary>
   //     public double Kerning { get => Properties.Attributes.Kerning; set => Properties.Attributes.Kerning = value; }

   //     public double Spacing { get => Properties.Attributes.Spacing; set => Properties.Attributes.Spacing = value; }

   //     /// <summary>
   //     /// FontStrike-out text
   //     /// </summary>
   //     public eStrikeType FontStrike { get => Properties.Attributes.FontStrike; set => Properties.Attributes.FontStrike = value; }

   //     /// <summary>
   //     /// Fontsize
   //     /// Spans from 0-4000
   //     /// </summary>
   //     public double FontSize { get => Properties.Attributes.FontSize; set => Properties.Attributes.FontSize = (float)value; }

   //     /// <summary>
   //     /// Underlined text
   //     /// </summary>
   //     public eUnderLineType FontUnderLine { get => Properties.Attributes.FontUnderLine; set => Properties.Attributes.FontUnderLine = value; }

   //     #endregion Attributes



   //     internal XmlElement PathElement
   //     {
   //         get
   //         {
   //             var node = (XmlElement)GetNode(_path);
   //             if (node == null)
   //             {
   //                 return (XmlElement)CreateNode(_path);
   //             }
   //             else
   //             {
   //                 return node;
   //             }
   //         }
   //     }
   //}
}
