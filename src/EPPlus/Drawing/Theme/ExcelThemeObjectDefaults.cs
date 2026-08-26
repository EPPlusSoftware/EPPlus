using OfficeOpenXml.Drawing.Shape;
using System.Xml;

namespace OfficeOpenXml.Drawing.Theme
{
    internal class ExcelThemeObjectDefaults : XmlHelper
    {
        private readonly ExcelThemeBase _theme;
        private readonly string _path = "objectDefaults";
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
                    _spDef = new DefaultShapeDefinition(NameSpaceManager, TopNode, _path +"\\spDef", _theme);
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
                    _lnDef = new DefaultShapeDefinition(NameSpaceManager, TopNode, _path + "\\lnDef", _theme);
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
                    _txDef = new DefaultShapeDefinition(NameSpaceManager, TopNode, _path + "\\txDef", _theme);
                }

                return _txDef;
            }
        }

        //TODO: Implement ExtLst
    }
}
