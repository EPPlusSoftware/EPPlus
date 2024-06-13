using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Layout settings
    /// </summary>
    public class ExcelLayout : XmlHelper
    {
        //Class for ExtLst Properties for later

        /// <summary>
        /// Manual layout settings for precise control of element position
        /// </summary>
        public ExcelManualLayout ManualLayout { get; }

        internal ExcelLayout(XmlNamespaceManager ns, XmlNode topNode, string path, string extLstPath, string[] schemaNodeOrder = null) : base(ns, topNode)
        {
            ManualLayout = new ExcelManualLayout(ns, topNode, $"{path}/c:manualLayout", $"{extLstPath}/c:manualLayout", schemaNodeOrder);
        }
    }
}
