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
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Datalabel on chart level. 
    /// This class is inherited by ExcelChartSerieDataLabel
    /// </summary>
    public class ExcelChartDataLabel : XmlHelper, IDrawingStyle
    {
        internal protected ExcelChartBase _chart;
        string _nodeName;
        internal ExcelChartDataLabel(ExcelChartBase chart, XmlNamespaceManager ns, XmlNode node, string nodeName, string[] schemaNodeOrder)
           : base(ns,node)
       {
            _nodeName = nodeName;
            AddSchemaNodeOrder(schemaNodeOrder, new string[] { "idx", "spPr", "txPr", "dLblPos", "showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize", "separator", "showLeaderLines" }, new int[] { 0, schemaNodeOrder.Length });

            AddSchemaNodeOrder(SchemaNodeOrder, ExcelDrawing._schemaNodeOrderSpPr);
            var fullNodeName = "c:" + nodeName;
            var topNode = GetNode(fullNodeName);
            if (topNode == null)
            {
                topNode = CreateNode(fullNodeName);
                topNode.InnerXml = "<c:showLegendKey val=\"0\" /><c:showVal val=\"0\" /><c:showCatName val=\"0\" /><c:showSerName val=\"0\" /><c:showPercent val=\"0\" /><c:showBubbleSize val=\"0\" /> <c:separator>\r\n</c:separator><c:showLeaderLines val=\"0\" />";                      
            }
            TopNode = topNode;
            _chart = chart;            
        }
        #region "Public properties"
        const string positionPath = "c:dLblPos/@val";
        /// <summary>
        /// Position of the labels
        /// </summary>
        public eLabelPosition Position
        {
            get
            {
                return GetPosEnum(GetXmlNodeString(positionPath));
            }
            set
            {
                if (ForbiddDataLabelPosition(_chart))
                {
                    throw (new InvalidOperationException("Can't set data label position on a 3D-chart"));
                }
                SetXmlNodeString(positionPath, GetPosText(value));
            }
        }

        internal static bool ForbiddDataLabelPosition(ExcelChartBase _chart)
        {
            return (_chart.IsType3D() && !_chart.IsTypePie() && _chart.ChartType != eChartType.Line3D)
                               || _chart.IsTypeDoughnut();
        }

        const string showValPath = "c:showVal/@val";
       /// <summary>
       /// Show the values 
       /// </summary>
        public bool ShowValue
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
        public bool ShowCategory
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
        public bool ShowSeriesName
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
        public bool ShowPercent
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
        public bool ShowLeaderLines
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
       public bool ShowBubbleSize
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
        public bool ShowLegendKey
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
        public string Separator
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
            
       ExcelDrawingFill _fill = null;
       /// <summary>
       /// Access fill properties
       /// </summary>
       public ExcelDrawingFill Fill
       {
           get
           {
               if (_fill == null)
               {
                   _fill = new ExcelDrawingFill(_chart, NameSpaceManager, TopNode, "c:spPr", SchemaNodeOrder);
               }
               return _fill;
           }
       }
       ExcelDrawingBorder _border = null;
       /// <summary>
       /// Access border properties
       /// </summary>
       public ExcelDrawingBorder Border
       {
           get
           {
               if (_border == null)
               {
                   _border = new ExcelDrawingBorder(_chart, NameSpaceManager, TopNode, "c:spPr/a:ln", SchemaNodeOrder);
               }
               return _border;
           }
       }
        ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// Effects
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (_effect == null)
                {
                    _effect = new ExcelDrawingEffectStyle(_chart, NameSpaceManager, TopNode, "c:spPr/a:effectLst", SchemaNodeOrder);
                }
                return _effect;
            }
        }
        ExcelDrawing3D _threeD = null;
        /// <summary>
        /// 3D properties
        /// </summary>
        public ExcelDrawing3D ThreeD
        {
            get
            {
                if (_threeD == null)
                {
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, "c:spPr", SchemaNodeOrder);
                }
                return _threeD;
            }
        }

        ExcelTextFont _font = null;
       /// <summary>
       /// Access font properties
       /// </summary>
       public ExcelTextFont Font
       {
           get
           {
               if (_font == null)
               {
                   _font = new ExcelTextFont(_chart, NameSpaceManager, TopNode, "c:txPr/a:p/a:pPr/a:defRPr", SchemaNodeOrder, CreateDefaultText);
               }
               return _font;
           }
       }
        void IDrawingStyleBase.CreatespPr()
        {
            CreatespPrNode();
        }

        private void CreateDefaultText()
        {
            if (TopNode.SelectSingleNode("c:txPr", NameSpaceManager) == null)
            {
                if (!ExistNode("c:spPr"))
                {
                    var spNode = CreateNode("c:spPr");
                    spNode.InnerXml = "<a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/>";
                }
                var node = CreateNode("c:txPr");
                node.InnerXml = "<a:bodyPr anchorCtr=\"1\" anchor=\"ctr\" bIns=\"19050\" rIns=\"38100\" tIns=\"19050\" lIns=\"38100\" wrap=\"square\" vert=\"horz\" vertOverflow=\"ellipsis\" spcFirstLastPara=\"1\" rot=\"0\"><a:spAutoFit/></a:bodyPr><a:lstStyle/>";
            }
            
        }

        ExcelTextBody _textBody = null;
        /// <summary>
        /// Access to text body properties
        /// </summary>
        public ExcelTextBody TextBody
        {
            get
            {
                if (_textBody == null)
                {
                    _textBody = new ExcelTextBody(NameSpaceManager, TopNode, "c:txPr/a:bodyPr", SchemaNodeOrder);
                }
                return _textBody;
            }
        }
        #endregion
        #region "Position Enum Translation"
        /// <summary>
        /// Translates the label position
        /// </summary>
        /// <param name="pos">The position enum</param>
        /// <returns>The string</returns>
        internal protected string GetPosText(eLabelPosition pos)
       {
           switch (pos)
           {
               case eLabelPosition.Bottom:
                   return "b";
               case eLabelPosition.Center:
                   return "ctr";
               case eLabelPosition.InBase:
                   return "inBase";
               case eLabelPosition.InEnd:
                   return "inEnd";
               case eLabelPosition.Left:
                   return "l";
               case eLabelPosition.Right:
                   return "r";
               case eLabelPosition.Top:
                   return "t";
               case eLabelPosition.OutEnd:
                   return "outEnd";
               default:
                   return "bestFit";
           }
       }
        /// <summary>
        /// Translates the enum position
        /// </summary>
        /// <param name="pos">The string value to translate</param>
        /// <returns>The enum value</returns>
       internal protected eLabelPosition GetPosEnum(string pos)
       {
           switch (pos)
           {
               case "b":
                   return eLabelPosition.Bottom;
               case "ctr":
                   return eLabelPosition.Center;
               case "inBase":
                   return eLabelPosition.InBase;
               case "inEnd":
                   return eLabelPosition.InEnd;
               case "l":
                   return eLabelPosition.Left;
               case "r":
                   return eLabelPosition.Right;
               case "t":
                   return eLabelPosition.Top;
               case "outEnd":
                   return eLabelPosition.OutEnd;
               default:
                   return eLabelPosition.BestFit;
           }
       }
    #endregion
    }
}
