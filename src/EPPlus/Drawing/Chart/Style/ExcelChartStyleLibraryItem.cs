/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.Style
{
    /// <summary>
    /// Represents a chart style xml document in the style library
    /// </summary>
    public class ExcelChartStyleLibraryItem
    {
        /// <summary>
        /// The id of the style
        /// </summary>
        public int Id { get; internal set; }
        /// <summary>
        /// The Xml as string
        /// </summary>
        public string XmlString { get; set; }
        XmlDocument _xmlDoc=null;
        /// <summary>
        /// The style xml document
        /// </summary>
        public XmlDocument XmlDocument
        {
            get
            {
                if(_xmlDoc==null)
                {
                    _xmlDoc = new XmlDocument();
                    XmlHelper.LoadXmlSafe(_xmlDoc, XmlString, Encoding.UTF8);
                }
                return _xmlDoc;
            }
        }
    }
}