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
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Utils.EnumUtils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Handles paragraph text
    /// </summary>
    public sealed class ExcelParagraph : ExcelTextFontRichText
    {
        internal ExcelParagraph(ExcelParagraphTextRunBase textRun) : 
            base(textRun)
        {
        }
        const string AligPath = "../../a:pPr/@algn";
        const string FldPath = "../a:fld";
        /// <summary>
        /// Text
        /// </summary>
        public eTextAlignment HorizontalAlignment
        {
            get
            {
                return _textRun.Paragraph.HorizontalAlignment;
            }
            set
            {
                _textRun.Paragraph.HorizontalAlignment = value;
            }
        }
        const string IndentLevelPath = "../../a:pPr/@lvl";
        /// <summary>
        /// Indent level for the paragraph. Ranges from 0-8;
        /// </summary>
        public int? IndentLevel
        {
            get
            {
                return _textRun.Paragraph.IndentLevel;
            }
            set
            {
                _textRun.Paragraph.IndentLevel = value ?? 0;
            }
        }

        /// <summary>
        /// Text
        /// </summary>
        public string Text
        {
            get
            {
               return _textRun.Text;
            }
            set
            {
                _textRun.Text = value;
            }
        }
        /// <summary>
        /// Text, adjusted for the Capitalization property
        /// </summary>
        public string DisplayedText
        {
            get
            {
                switch (_textRun.Capitalization)
                {

                    case eTextCapsType.All:
                        return _textRun.Text.ToUpper();
                    case eTextCapsType.Small:
                        return _textRun.Text.ToLower();
                    default:
                        return _textRun.Text;

                }
            }
        }

        /// <summary>
        /// If the paragraph is the first in the collection
        /// </summary>
        public bool IsFirstInParagraph
        {
            get
            {
                return _textRun.IsFirstInParagraph;
            }
        }
        /// <summary>
        /// If the paragraph is the last in the collection
        /// </summary>
        public bool IsLastInParagraph
        {
            get
            {
                return _textRun.IsLastInParagraph;
            }
        }

        internal bool IsInParagraph(ExcelDrawingParagraph paragraph)
        {
            return _textRun.Paragraph.Equals(paragraph);
        }
        internal bool IsTextRun(ExcelParagraphTextRunBase paragraph)
        {
            return _textRun.Equals(paragraph);
        }
    }
}
