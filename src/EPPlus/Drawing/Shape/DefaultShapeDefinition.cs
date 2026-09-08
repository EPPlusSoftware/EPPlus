using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Shape.Style;
using System.Xml;

namespace OfficeOpenXml.Drawing.Shape
{
    /// <summary>
    /// Roughly Represents CT_DefaultShapeDefinition
    /// </summary>
    internal class DefaultShapeDefinition : XmlHelper
    {
        string _fillPath = "{0}/{1}:spPr";
        string _defaultTextBodyPath = "{0}/{1}:bodyPr";
        string _stylePath = "{0}/{1}:style";

        //Do we support this? Does Excel? Excel appears to in this specific case.
        //TODO: ImplementTextList

        //TODO: Implement ExtLst
        //private ExtLst

        private readonly IPictureRelationDocument _pictureRelationDocument;

        string _prefix;

        internal DefaultShapeDefinition(XmlNamespaceManager nsm, XmlNode topNode, string path, IPictureRelationDocument pictureRelationDocument, string prefix = "a") : base(nsm, topNode)
        {
            _prefix = prefix;

            _fillPath = string.Format(_fillPath, path, _prefix);
            _defaultTextBodyPath = string.Format(_defaultTextBodyPath, path, _prefix);
            _stylePath = string.Format(_stylePath, path, _prefix);
        }


        private ExcelDrawingFill _fill;
        /// <remarks/>
        /// <summary>
        /// Reference to fill settings for a chart part
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(_pictureRelationDocument, NameSpaceManager, TopNode, _fillPath, SchemaNodeOrder);
                }
                return _fill;
            }
        }

        private ExcelTextBody _defaultTextBody = null;
        /// <summary>
        /// Reference to default text body run settings for a chart part
        /// </summary>
        public ExcelTextBody DefaultTextBody
        {
            get
            {
                if (_defaultTextBody == null)
                {
                    _defaultTextBody = new ExcelTextBody(_pictureRelationDocument, NameSpaceManager, TopNode, _defaultTextBodyPath);
                }
                return _defaultTextBody;

            }
        }

        ExcelShapeStyleEntry _style;

        /// <summary>
        /// Reference to default text body run settings for a chart part
        /// </summary>
        public ExcelShapeStyleEntry Style
        {
            get
            {
                if (_style == null)
                {
                    _style = new ExcelShapeStyleEntry(NameSpaceManager, TopNode, _stylePath, _pictureRelationDocument, _prefix);
                }
                return _style;

            }
        }
    }
}
