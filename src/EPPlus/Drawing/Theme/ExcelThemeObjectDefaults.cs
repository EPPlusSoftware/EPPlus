using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Theme
{
    internal class ExcelThemeObjectDefaults : XmlHelper
    {
        private readonly ExcelThemeBase _theme;

        public ExcelThemeObjectDefaults(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelThemeBase theme) : base(nameSpaceManager, topNode)
        {
            _theme = theme;
        }
    }
}
