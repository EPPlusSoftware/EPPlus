/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    11/24/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    /// <summary>
    /// An abstract class used for formcontrols with text properties.
    /// </summary>
    public abstract class ExcelControlWithText : ExcelControl
    {
        private string _paragraphPath = "xdr:sp/xdr:txBody/a:p";
        private string _lockTextPath = "xdr:sp/@fLocksText";

        internal ExcelControlWithText(ExcelDrawings drawings, XmlNode drawingNode, ControlInternal control, ZipPackagePart part, XmlDocument ctrlPropXml, ExcelGroupShape parent = null) : 
            base(drawings, drawingNode, control, part, ctrlPropXml, parent)
        {

        }

        internal ExcelControlWithText(ExcelDrawings drawings, XmlElement drawNode, string name, ExcelGroupShape parent=null) : 
            base(drawings, drawNode, name, parent)
        {
        }

        /// <summary>
        /// Text inside the shape
        /// </summary>
        public string Text
        {
            get
            {
                return RichText.Text;
            }
            set
            {
                if(RichText.Count==1)
                {
                    RichText[0].Text = value;
                }
                else
                {
                    RichText.Clear();
                    RichText.Text = value;
                }
                _vml.Text = value;
            }
        }
        ExcelParagraphCollection _richText = null;

        /// <summary>
        /// Richtext collection. Used to format specific parts of the text
        /// </summary>
        public ExcelParagraphCollection RichText
        {
            get
            {
                if (_richText == null)
                {
                    _richText = new ExcelParagraphCollection(this, NameSpaceManager, TopNode, _paragraphPath, SchemaNodeOrder);
                }
                return _richText;
            }
        }
        /// <summary>
        /// Gets or sets whether a controls text is locked when the worksheet is protected.
        /// </summary>
        public bool LockedText
        {
            get
            {
                return _ctrlProp.GetXmlNodeBool("@lockText");
            }
            set
            {
                _ctrlProp.SetXmlNodeBool("@lockText", value);
                SetXmlNodeBool(_lockTextPath, value);
            }
        }
        ExcelTextBody _textBody = null;
        /// <summary>
        /// Access to text body properties.
        /// </summary>
        public ExcelTextBody TextBody
        {
            get
            {
                if (_textBody == null)
                {
                    _textBody = new ExcelTextBody(NameSpaceManager, TopNode, "xdr:sp/xdr:txBody/a:bodyPr", this.SchemaNodeOrder);
                }
                return _textBody;
            }
        }
    }
}
