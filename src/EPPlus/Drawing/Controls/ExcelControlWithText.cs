using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    public abstract class ExcelControlWithText : ExcelControl
    {
        private string _paragraphPath = "xdr:sp/xdr:txBody/a:p";
        private string _lockTextPath = "xdr:sp/@fLocksText";

        internal ExcelControlWithText(ExcelDrawings drawings, XmlNode drawingNode, ControlInternal control, ZipPackageRelationship rel, XmlDocument ctrlPropXml, ExcelGroupShape parent = null) : 
            base(drawings, drawingNode, control, rel, ctrlPropXml, parent)
        {

        }

        protected ExcelControlWithText(ExcelDrawings drawings, XmlElement drawNode) : base(drawings, drawNode)
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
                RichText.Clear();
                RichText.Text = value;
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
    }
}
