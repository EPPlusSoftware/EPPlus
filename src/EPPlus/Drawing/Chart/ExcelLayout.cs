using System;
using System.Collections.Generic;
using System.IO;
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
        string _path = null;
        internal ExcelLayout(XmlNamespaceManager ns, XmlNode topNode, string path, string extLstPath, string[] schemaNodeOrder = null) : base(ns, topNode)
        {
            _path = path;
            ManualLayout = new ExcelManualLayout(ns, topNode, $"{path}/c:manualLayout", $"{extLstPath}/c:manualLayout", schemaNodeOrder);
        }
        internal bool HasLayout
        {
            get
            {
                var n = GetNode(_path);
                return n != null && (n.Attributes.Count > 0 || n.ChildNodes.Count > 0);
            }
        }
    }
}
