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
        private const string defaultSpDefInnerXml = "<a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"15000\" /></a:schemeClr></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\" /></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\" /></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\" /></a:fontRef></a:style>";

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
                    var spDefNode = TopNode.SelectSingleNode(_spDefPath, NameSpaceManager);
                    //Despite there being no SpDef node/no child nodes Excel Acts as if the @defaultSpDefXml is there.
                    //Therefore if the node is not there or if it is empty create the default 
                    if (spDefNode == null || spDefNode.HasChildNodes == false)
                    {
                        spDefNode = CreateNode(_spDefPath);
                        spDefNode.InnerXml = defaultSpDefInnerXml;
                    }
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
