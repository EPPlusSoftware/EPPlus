using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Interfaces;
using System.Xml;

namespace OfficeOpenXml.Drawing.Shape.Style
{
    //Style (CT_ShapeStyle node)
    internal class ExcelShapeStyleEntry : XmlHelper
    {
        string _fillReferencePath = "{0}/{1}:fillRef";
        string _borderReferencePath = "{0}/{1}:lnRef";
        string _effectReferencePath = "{0}/{1}:effectRef";
        string _fontReferencePath = "{0}/{1}:fontRef";

        private readonly IPictureRelationDocument _pictureRelationDocument;
        internal ExcelShapeStyleEntry(XmlNamespaceManager nsm, XmlNode topNode, string path, IPictureRelationDocument pictureRelationDocument, string prefix = "a") : base(nsm, topNode)
        {

        }
        private ShapeStyleReference _borderReference = null;
        /// Border reference. 
        /// Contains an index reference to the theme and a color to be used in border styling
        public ShapeStyleReference BorderReference
        {
            get
            {
                if (_borderReference == null)
                {
                    _borderReference = new ShapeStyleReference(NameSpaceManager, TopNode, _borderReferencePath);
                }
                return _borderReference;
            }
        }
        private ShapeStyleReference _fillReference = null;
        /// <summary>
        /// Fill reference. 
        /// Contains an index reference to the theme and a fill color to be used in fills
        /// </summary>
        public ShapeStyleReference FillReference
        {
            get
            {
                if (_fillReference == null)
                {
                    _fillReference = new ShapeStyleReference(NameSpaceManager, TopNode, _fillReferencePath);
                }
                return _fillReference;
            }
        }
        private ShapeStyleReference _effectReference = null;
        /// <summary>
        /// Effect reference. 
        /// Contains an index reference to the theme and a color to be used in effects
        /// </summary>
        public ShapeStyleReference EffectReference
        {
            get
            {
                if (_effectReference == null)
                {
                    _effectReference = new ShapeStyleReference(NameSpaceManager, TopNode, _effectReferencePath);
                }
                return _effectReference;
            }
        }
        ExcelChartStyleFontReference _fontReference = null;
        /// <summary>
        /// Font reference. 
        /// Contains an index reference to the theme and a color to be used for font styling
        /// </summary>
        public ExcelChartStyleFontReference FontReference
        {
            get
            {
                if (_fontReference == null)
                {
                    _fontReference = new ExcelChartStyleFontReference(NameSpaceManager, TopNode, _fontReferencePath);
                }
                return _fontReference;
            }
        }
    }
}
