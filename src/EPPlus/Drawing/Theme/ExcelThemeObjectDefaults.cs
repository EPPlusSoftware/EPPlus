using OfficeOpenXml.Drawing.Shape;
using System.Xml;

namespace OfficeOpenXml.Drawing.Theme
{
    internal class ExcelThemeObjectDefaults : XmlHelper
    {
        private readonly ExcelThemeBase _theme;
        private readonly string _path = "objectDefaults";
        private readonly string _spDefPath = "a:spDef";
        private readonly string _lnDefPath = "a:lnDef";
        private readonly string _txDefPath = "a:txDef";
        public ExcelThemeObjectDefaults(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelThemeBase theme) : base(nameSpaceManager, topNode)
        {
            _theme = theme;
        }

        DefaultShapeDefinition _spDef = null;
        DefaultShapeDefinition _lnDef = null;
        DefaultShapeDefinition _txDef = null;

        public DefaultShapeDefinition ShapeDefinition
        {
            get
            {
                if (_spDef == null)
                {
                    var test = TopNode.SelectSingleNode(_spDefPath, NameSpaceManager);
                    _spDef = new DefaultShapeDefinition(NameSpaceManager, TopNode, _spDefPath, _theme);
                }

                return _spDef;
            }
        }

        public DefaultShapeDefinition LineDefinition
        {
            get
            {
                if (_lnDef == null)
                {
                    _lnDef = new DefaultShapeDefinition(NameSpaceManager, TopNode, _lnDefPath, _theme);
                }

                return _lnDef;
            }
        }

        public DefaultShapeDefinition TextDefinition
        {
            get
            {
                if (_txDef == null)
                {
                    _txDef = new DefaultShapeDefinition(NameSpaceManager, TopNode, _txDefPath, _theme);
                }

                return _txDef;
            }
        }

        //TODO: Implement ExtLst
    }
}
