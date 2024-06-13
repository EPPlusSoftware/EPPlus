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
using System;
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml.Drawing.Chart.DataLabling;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Settings for a charts data lables
    /// </summary>
    public class ExcelChartDataLabelStandard : ExcelChartDataLabel
    {
        Guid _guidId;

        internal ExcelChartDataLabelStandard(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, string nodeName, string[] schemaNodeOrder)
           : base(chart, ns, node, nodeName, "c")
        {
            AddSchemaNodeOrder([""], LabelNodeHolder.DataLabels.NodeOrder);
            var order = SchemaNodeOrder;

            if (nodeName == "dLbl" || nodeName == "")
            {
                AddSchemaNodeOrder([""], LabelNodeHolder.DataLabel.NodeOrder);
                
                TopNode = node;
                
                var extPath = "c:extLst/c:ext";

                NameSpaceManager.AddNamespace("xmlns:c15", ExcelPackage.schemaChart2012);
                NameSpaceManager.AddNamespace("c16", ExcelPackage.schemaChart2014);

                XmlElement el = (XmlElement)CreateNode($"{extPath}");
                el.SetAttribute("xmlns:c15", ExcelPackage.schemaChart2012);
                SetXmlNodeString($"{extPath}/@uri","{CE6537A1-D6FC-4f65-9D91-7224C49458BB}");

                XmlElement element = (XmlElement)CreateNode($"{extPath}", false, true);
                element.SetAttribute("xmlns:c16", ExcelPackage.schemaChart2014);
                SetXmlNodeString($"{extPath}[2]/@uri", "{C3380CC4-5D6E-409C-BE32-E72D297353CC}");
                //SetXmlNodeString($"{extPath}[2]", "{C3380CC4-5D6E-409C-BE32-E72D297353CC}");

                _guidId = Guid.NewGuid();

                var extNode2 = GetNode($"{extPath}[2]");
                var uniqueIdNode = (XmlElement)CreateNode(extNode2, "c16:uniqueID");
                uniqueIdNode.SetAttribute("val", $"{{{_guidId}}}");
                //XmlElement idElement = (XmlElement)CreateNode($"{extPath}[2]", false, true);
                //SetXmlNodeString($"{extPath}[2][1]/c16:uniqueId/@val", $"{{{_guidId}}}");


                // //SetXmlNodeString($"{extPath}/@uri", "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}");

                //var aNode = GetNode($"{extPath}");

                //var item = aNode.Attributes.Item(0);
                // var attr = aNode.OwnerDocument.CreateAttribute("c15", "xmlns", "http://schemas.microsoft.com/office/drawing/2012/chart");
                // //attr.Value = "http://schemas.microsoft.com/office/drawing/2012/chart";
                // //aNode.Attributes.Prepend(attr);

                // var element = aNode.OwnerDocument.CreateElement("c","ext", "http://schemas.microsoft.com/office/drawing/2014/chart");
                // //var attr2 = aNode.OwnerDocument.CreateAttribute("c16", "xmlns", "http://schemas.microsoft.com/office/drawing/2014/chart");
                //// attr2.Value = "http://schemas.microsoft.com/office/drawing/2014/chart";
                // //element.Attributes.Append(attr2);
                // var uri2 = aNode.OwnerDocument.CreateAttribute("uri");
                // uri2.Value = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}";
                // element.Attributes.Append(uri2);

                // aNode.ParentNode.AppendChild(element);

                // var idElement = aNode.OwnerDocument.CreateElement("c16", "uniqueId", null);
                // var idAttr = aNode.OwnerDocument.CreateAttribute("val");
                // idAttr.Value = $"{{{Guid.NewGuid()}}}";
                // idElement.Attributes.Append(idAttr);
                // element.AppendChild(idElement);

                // var checkNode = aNode;
                //SetXmlNodeString($"{extPath}/@xmlns:c15", "http://schemas.microsoft.com/office/drawing/2012/chart");
            }
            else
            {
                var fullNodeName = "c:" + nodeName;
                var topNode = GetNode(fullNodeName);
                if (topNode == null)
                {
                    topNode = CreateNode(fullNodeName);
                    topNode.InnerXml = "<c:showLegendKey val=\"0\" /><c:showVal val=\"0\" /><c:showCatName val=\"0\" /><c:showSerName val=\"0\" /><c:showPercent val=\"0\" /><c:showBubbleSize val=\"0\" /> <c:separator>\r\n</c:separator><c:showLeaderLines val=\"0\" />";
                }
                TopNode = topNode;
            }
        }

        const string positionPath = "c:dLblPos/@val";
        /// <summary>
        /// Position of the labels
        /// Note: Only Center, InEnd and InBase are allowed for dataLabels on stacked columns 
        /// </summary>
        public override eLabelPosition Position
        {
            get
            {
                return GetPosEnum(GetXmlNodeString(positionPath));
            }
            set
            {
                if (ForbiddDataLabelPosition(_chart))
                {
                    throw new InvalidOperationException("Can't set data label position on a 3D-chart");
                }
                SetXmlNodeString(positionPath, GetPosText(value));
            }
        }
        internal static bool ForbiddDataLabelPosition(ExcelChart _chart)
        {
            return _chart.IsType3D() && !_chart.IsTypePie() && _chart.ChartType != eChartType.Line3D
                               || _chart.IsTypeDoughnut();
        }
        const string showValPath = "c:showVal/@val";
        /// <summary>
        /// Show the values 
        /// </summary>
        public override bool ShowValue
        {
            get
            {
                return GetXmlNodeBool(showValPath);
            }
            set
            {
                SetXmlNodeString(showValPath, value ? "1" : "0");
            }
        }
        const string showCatPath = "c:showCatName/@val";
        /// <summary>
        /// Show category names  
        /// </summary>
        public override bool ShowCategory
        {
            get
            {
                return GetXmlNodeBool(showCatPath);
            }
            set
            {
                SetXmlNodeString(showCatPath, value ? "1" : "0");
            }
        }
        const string showSerPath = "c:showSerName/@val";
        /// <summary>
        /// Show series names
        /// </summary>
        public override bool ShowSeriesName
        {
            get
            {
                return GetXmlNodeBool(showSerPath);
            }
            set
            {
                SetXmlNodeString(showSerPath, value ? "1" : "0");
            }
        }
        const string showPerentPath = "c:showPercent/@val";
        /// <summary>
        /// Show percent values
        /// </summary>
        public override bool ShowPercent
        {
            get
            {
                return GetXmlNodeBool(showPerentPath);
            }
            set
            {
                SetXmlNodeString(showPerentPath, value ? "1" : "0");
            }
        }
        const string showLeaderLinesPath = "c:showLeaderLines/@val";
        /// <summary>
        /// Show the leader lines
        /// </summary>
        public override bool ShowLeaderLines
        {
            get
            {
                return GetXmlNodeBool(showLeaderLinesPath);
            }
            set
            {
                SetXmlNodeString(showLeaderLinesPath, value ? "1" : "0");
            }
        }
        const string showBubbleSizePath = "c:showBubbleSize/@val";
        /// <summary>
        /// Show Bubble Size
        /// </summary>
        public override bool ShowBubbleSize
        {
            get
            {
                return GetXmlNodeBool(showBubbleSizePath);
            }
            set
            {
                SetXmlNodeString(showBubbleSizePath, value ? "1" : "0");
            }
        }
        const string showLegendKeyPath = "c:showLegendKey/@val";
        /// <summary>
        /// Show the Lengend Key
        /// </summary>
        public override bool ShowLegendKey
        {
            get
            {
                return GetXmlNodeBool(showLegendKeyPath);
            }
            set
            {
                SetXmlNodeString(showLegendKeyPath, value ? "1" : "0");
            }
        }
        const string separatorPath = "c:separator";
        /// <summary>
        /// Separator string 
        /// </summary>
        public override string Separator
        {
            get
            {
                return GetXmlNodeString(separatorPath);
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    DeleteNode(separatorPath);
                }
                else
                {
                    SetXmlNodeString(separatorPath, value);
                }
            }
        }
    }
}
