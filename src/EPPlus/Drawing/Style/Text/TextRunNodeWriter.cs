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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;
using OfficeOpenXml.WritingExtension;
namespace OfficeOpenXml.Style
{

    /// <summary>
    /// Used by Rich-text and Paragraphs.
    /// </summary>
    public class TextRunNodeWriter : XmlHelper
    {
        string _path;
        internal XmlNode _rootNode;

        internal TextRunNodeWriter(XmlNamespaceManager namespaceManager, XmlNode rootNode, string path, string[] schemaNodeOrder)
            : base(namespaceManager, rootNode)
        {
            AddSchemaNodeOrder(schemaNodeOrder, new string[] { "bodyPr", "lstStyle", "p", "pPr", "defRPr", "solidFill", "highlight", "uFill", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "r", "rPr", "t" });
            _rootNode = rootNode;
            if (path != "")
            {
                XmlNode node = rootNode.SelectSingleNode(path, namespaceManager);
                if (node != null)
                {
                    TopNode = node;
                }
            }
            _path = path;
        }

        const string _underLineColorPath = "a:uFill/a:solidFill/a:srgbClr/@val";

        static string fontString = "a:{0}/@typeface";
        string _fontLatinPath = string.Format(fontString, "latin");
        string _fontEaPath = string.Format(fontString, "ea");
        string _fontCsPath = string.Format(fontString, "cs");
        string _fontSymPath = string.Format(fontString, "cs");

        string _rtlPath = "/a:rtl/@w:val";
        string _boldPath = "@b";
        string _italicPath = "@i";


        //Drawing fill has its own read/write handling that is set directly. Therefore no need for it here (?)

        //TODO: The same for Reading/Using the new Parse functions in ExcelTextFont
        //Arguably this is just a broken out fancy setter. Perhaps it would be better to use the existing class?
        //Simply intercept all the setter and getters to the new file, read in the entire file/settings in constructor
        //Call and write all xml on save? That would slow down save operation instead of doing it continously though...
        //Though with this setup we could have a property and simply run individual saves for each prop via a global setting or something.

        internal void WriteXml(RegularTextRun txtRun)
        {
            WriteUnderLineColor(txtRun.UnderLineColor);

            txtRun.LatinFont?.TryAct(x => WriteFont(_fontLatinPath, x));
            txtRun.EastAsianFont?.TryAct(x => WriteFont(_fontEaPath, x));
            txtRun.ComplexFont?.TryAct(x => WriteFont(_fontCsPath, x));
            txtRun.SymbolFont?.TryAct(x => WriteFont(_fontSymPath, x));

            WriteBool(_rtlPath, txtRun.rtl);
            WriteBool(_boldPath, txtRun.Bold);
            WriteBool(_italicPath, txtRun.Italic);
            WriteUnderline(txtRun.UnderLine);
            WriteStrike(txtRun.Strike);

            WriteSize((float)txtRun.FontSize);
            WriteKerning(txtRun.Kerning);
            WriteCaps(txtRun.Capitalization);
            WriteBaseLine(txtRun.Baseline);

            txtRun.Text?.TryAct(WriteInnerText);
        }

        private void WriteUnderLineColor(Color value)
        {
            SetXmlNodeString(_underLineColorPath, value.ToArgb().ToString("X").Substring(2, 6));
        }

        private void WriteFont(string fontPath, string content)
        {
            SetXmlNodeString(fontPath, content);
        }
        private void WriteBool(string path, bool boolValue)
        {
            SetXmlNodeString(path, boolValue ? "1" : "0");
        }

        string _underLinePath = "@u";
        private void WriteUnderline(eUnderLineType underlineType)
        {
            SetXmlNodeString(_underLinePath, underlineType.TranslateUnderlineText());
        }

        string _strikePath = "@strike";
        private void WriteStrike(eStrikeType strikeType)
        {
            SetXmlNodeString(_strikePath, strikeType.TranslateStrikeTypeText());
        }

        string _sizePath = "@sz";
        private void WriteSize(float size)
        {
            SetXmlNodeString(_sizePath, ((int)(size * 100)).ToString());
        }

        string _kernPath = "@kern";
        private void WriteKerning(double kernSize)
        {
            SetXmlNodeFontSize(_kernPath, kernSize, "Kerning");
        }

        private void WriteCaps(eTextCapsType capsType)
        {
            SetXmlNodeString($"{_path}/@cap", capsType.ToEnumString());
        }

        private void WriteBaseLine(double baseLine)
        {
            SetXmlNodePercentage($"{_path}/@baseline", baseLine);
        }

        string _innerTextPath = "../t";
        private void WriteInnerText(string innerString)
        {
            SetXmlNodeString(_innerTextPath, innerString);
        }
    }
}