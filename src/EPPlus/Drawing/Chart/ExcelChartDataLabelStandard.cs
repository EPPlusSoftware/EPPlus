﻿/*************************************************************************************************
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
        internal ExcelChartDataLabelStandard(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, string nodeName, string[] schemaNodeOrder)
           : base(chart, ns, node, nodeName, "c")
        {
            //var groupDLbls = 

            //var CT_DLblsNodeOrder = new List<string>() { "dLbl", };

            //var DLblShared = new string[] { "numFmt", "spPr", "txPr", "dLblPos", "showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize", "separator" };

            //var choiceNode = new string[] { "dLbl", "delete" };

            //var groupDLbls = Array.Copy(choiceNode);

            //AddSchemaNodeOrder([""], CTDataLabels.NodeOrder);

            SchemaNodeOrder = LabelNodeHolder.DataLabels.NodeOrder;

            var order = SchemaNodeOrder;

            //AddSchemaNodeOrder(choiceNode, DLblShared);
            //AddSchemaNodeOrder(SchemaNodeOrder, ["showLeaderLines", "leaderLines", "extLst"]);
            //AddSchemaNodeOrder(SchemaNodeOrder, DLblShared);

            if (nodeName == "dLbl" || nodeName == "")
            {
                SchemaNodeOrder = LabelNodeHolder.DataLabel.NodeOrder;
                //var dLblNodes = new string[] { "idx, delete" };
                //var groupDLbl = new string[] { "layout, tx" };
                //Array.Copy(DLblShared, groupDLbl, DLblShared.Length);

                //SchemaNodeOrder = dLblNodes;
                //AddSchemaNodeOrder(SchemaNodeOrder, groupDLbl);
                //AddSchemaNodeOrder(SchemaNodeOrder, ["extLst"]);

                //Adds
                //SchemaNodeOrder = 
                //groupDLbl = new string[] { dLblNodes[0], dLblNodes[1], groupDLbl[0], groupDLbl[1] };

                //var ctDLbl = string.Concat(dLblNodes, )
                //AddSchemaNodeOrder(dLblNodes,)
                TopNode = node;
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
