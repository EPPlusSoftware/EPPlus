/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// The visibility for various elements in the chart
    /// </summary>
    public class ExcelChartExSerieElementVisibilities : XmlHelper
    {
        const string _path = "cx:layout/cx:visibility";
        public ExcelChartExSerieElementVisibilities(XmlNamespaceManager nsm, XmlNode node, string[] schemaNodeOrder) : base(nsm, node)
        {
            SchemaNodeOrder = schemaNodeOrder;
        }
        /// <summary>
        /// The visibility of connector lines between data points
        /// </summary>
        public bool ConnectorLines 
        { 
            get
            {
                return GetXmlNodeBool($"{_path}/@connectorLines");
            }
            set
            {
                SetXmlNodeBool($"{_path}/@connectorLines", value);
            }
        }
        /// <summary>
        /// The visibility of connector lines between data points
        /// </summary>
        public bool MeanLine 
        {
            get
            {
                return GetXmlNodeBool($"{_path}/@meanLine");
            }
            set
            {
                SetXmlNodeBool($"{_path}/@meanLine", value);
            }
        }
        /// <summary>
        /// The visibility of markers denoting the mean
        /// </summary>
        public bool MeanMarker
        {
            get
            {
                return GetXmlNodeBool($"{_path}/@meanMarker");
            }
            set
            {
                SetXmlNodeBool($"{_path}/@meanMarker", value);
            }
        }
        /// <summary>
        /// The visibility of non-outlier data points
        /// </summary>
        public bool NonOutliers
        {
            get
            {
                return GetXmlNodeBool($"{_path}/@nonOutliers");
            }
            set
            {
                SetXmlNodeBool($"{_path}/@nonOutliers", value);
            }
        }
        /// <summary>
        /// The visibility of outlier data points
        /// </summary>
        public bool Outliers 
        {
            get
            {
                return GetXmlNodeBool($"{_path}/@outliers");
            }
            set
            {
                SetXmlNodeBool($"{_path}/@outliers", value);
            }
        }
    }
}
