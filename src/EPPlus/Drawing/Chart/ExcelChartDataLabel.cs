/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/24/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
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
    public abstract class ExcelChartDataLabel : XmlHelper, IDrawingStyle
    {
        internal protected ExcelChart _chart;
        internal protected string _nodeName;
        private string _nsPrefix;
        private readonly string _formatPath;
        private readonly string _sourceLinkedPath;

        internal ExcelChartDataLabel(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, string nodeName, string nsPrefix)
           : base(ns,node)
       {
            _nodeName = nodeName;
            _chart = chart;
            _nsPrefix = nsPrefix;
            _formatPath = $"{nsPrefix}:numFmt/@formatCode";
            _sourceLinkedPath = $"{nsPrefix}:numFmt/@sourceLinked";
        }
        #region "Public properties"
        public abstract eLabelPosition Position { get; set; }
       /// <summary>
       /// Show the values 
       /// </summary>
        public abstract bool ShowValue
        {
            get;
            set;
        }
       /// <summary>
       /// Show category names  
       /// </summary>
        public abstract bool ShowCategory
        {
            get;
            set;
        }
       /// <summary>
       /// Show series names
       /// </summary>
        public abstract bool ShowSeriesName
        {
            get;
            set;
        }
       /// <summary>
       /// Show percent values
       /// </summary>
        public abstract bool ShowPercent
        {
            get;
            set;
        }
       /// <summary>
       /// Show the leader lines
       /// </summary>
        public abstract bool ShowLeaderLines
        {
           get;
           set;
       }
       /// <summary>
       /// Show Bubble Size
       /// </summary>
       public abstract bool ShowBubbleSize
       {
            get;
            set;
       }
        /// <summary>
        /// Show the Lengend Key
        /// </summary>
        public abstract bool ShowLegendKey
        {
            get;
            set;
       }
       /// <summary>
       /// Separator string 
       /// </summary>
        public abstract string Separator
       {
            get;
            set;
       }

        /// <summary>
        /// The Numberformat string.
        /// </summary>
        public string Format
        {
            get
            {
                return GetXmlNodeString(_formatPath);
            }
            set
            {
                SetXmlNodeString(_formatPath, value);
                if (string.IsNullOrEmpty(value))
                {
                    SourceLinked = true;
                }
                else
                {
                    SourceLinked = false;
                }
            }
        }
        /// <summary>
        /// The Numberformats are linked to the source data.
        /// </summary>
        public bool SourceLinked
        {
            get
            {
                return GetXmlNodeBool(_sourceLinkedPath);
            }
            set
            {
                SetXmlNodeBool(_sourceLinkedPath, value);
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
                   _fill = new ExcelDrawingFill(_chart, NameSpaceManager, TopNode, $"{_nsPrefix}:spPr", SchemaNodeOrder);
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
                   _border = new ExcelDrawingBorder(_chart, NameSpaceManager, TopNode, $"{_nsPrefix}:spPr/a:ln", SchemaNodeOrder);
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
                    _effect = new ExcelDrawingEffectStyle(_chart, NameSpaceManager, TopNode, $"{_nsPrefix}:spPr/a:effectLst", SchemaNodeOrder);
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
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, $"{_nsPrefix}:spPr", SchemaNodeOrder);
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
                   _font = new ExcelTextFont(_chart, NameSpaceManager, TopNode, $"{_nsPrefix}:txPr/a:p/a:pPr/a:defRPr", SchemaNodeOrder, CreateDefaultText);
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
            if (TopNode.SelectSingleNode($"{_nsPrefix}:txPr", NameSpaceManager) == null)
            {
                if (!ExistsNode($"{_nsPrefix}:spPr"))
                {
                    var spNode = CreateNode($"{_nsPrefix}:spPr");
                    spNode.InnerXml = "<a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/>";
                }
                var node = CreateNode($"{_nsPrefix}:txPr");
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
                    _textBody = new ExcelTextBody(NameSpaceManager, TopNode, $"{_nsPrefix}:txPr/a:bodyPr", SchemaNodeOrder);
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
