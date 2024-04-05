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
using System.Xml;
using OfficeOpenXml.Style;
using System.Globalization;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Utils.Extensions;
using System.Runtime.InteropServices;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// An axis for a chart
    /// </summary>
    public abstract class ExcelChartAxis : XmlHelper, IDrawingStyle, IStyleMandatoryProperties
    {
        /// <summary>
        /// Type of axis
        /// </summary>
        internal ExcelChart _chart;
        internal string _nsPrefix;
        private readonly string _minorGridlinesPath;
        private readonly string _majorGridlinesPath;
        private readonly string _formatPath;
        private readonly string _sourceLinkedPath;

        internal ExcelChartAxis(ExcelChart chart, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string nsPrefix) :
            base(nameSpaceManager, topNode)
        {
            _chart = chart;
            _nsPrefix = nsPrefix;
            _formatPath = $"{_nsPrefix}:numFmt/@formatCode";
            _sourceLinkedPath = $"{_nsPrefix}:numFmt/@sourceLinked";
            _minorGridlinesPath = $"{nsPrefix}:minorGridlines";
            _majorGridlinesPath = $"{nsPrefix}:majorGridlines";
        }
        internal abstract string Id
        {
            get;
        }
        /// <summary>
        /// Get or Sets the major tick marks for the axis. 
        /// </summary>
        public abstract eAxisTickMark MajorTickMark
        {
            get;
            set;
        }

        /// <summary>
        /// Get or Sets the minor tick marks for the axis. 
        /// </summary>
        public abstract eAxisTickMark MinorTickMark
        {
            get;
            set;
        }
        /// <summary>
        /// The type of axis
        /// </summary>
        internal abstract eAxisType AxisType
        {
            get;
        }
        /// <summary>
        /// Where the axis is located
        /// </summary>
        public abstract eAxisPosition AxisPosition
        {
            get;
            internal set;
        }
        /// <summary>
        /// Where the axis crosses
        /// </summary>
        public abstract eCrosses Crosses
        {
            get;
            set;
        }
        /// <summary>
        /// How the axis are crossed
        /// </summary>
        public abstract eCrossBetween CrossBetween
        {
            get;
            set;
        }
        /// <summary>
        /// The value where the axis cross. 
        /// Null is automatic
        /// </summary>
        public abstract double? CrossesAt
        {
            get;
            set;
        }
        /// <summary>
        /// The Numberformat used
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
        /// <summary>
                 /// The Position of the labels
                 /// </summary>
        public abstract eTickLabelPosition LabelPosition
        {
            get;
            set;
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Access to fill properties
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
        /// Access to border properties
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
        /// Access to font properties
        /// </summary>
        public ExcelTextFont Font
        {
            get
            {
                if (_font == null)
                {
                    _font = new ExcelTextFont(_chart, NameSpaceManager, TopNode, $"{_nsPrefix}:txPr/a:p/a:pPr/a:defRPr", SchemaNodeOrder);
                }
                return _font;
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
                    _textBody = new ExcelTextBody(NameSpaceManager, TopNode, $"{_nsPrefix}:txPr/a:bodyPr", SchemaNodeOrder, Font.CreateTopNode);
                }
                return _textBody;
            }
        }
		ExcelDrawingTextSettings _textSettings = null;
		/// <summary>
		/// String settings like fills, text outlines and effects 
		/// </summary>
		public ExcelDrawingTextSettings TextSettings
		{
			get
			{
				if (_textSettings == null)
				{
					_textSettings = new ExcelDrawingTextSettings(_chart, NameSpaceManager, TopNode, $"{_nsPrefix}:txPr/a:p/a:pPr/a:defRPr", SchemaNodeOrder);
				}
				return _textSettings;
			}
		}

		void IDrawingStyleBase.CreatespPr()
        {
            CreatespPrNode($"{_nsPrefix}:spPr");
        }

        /// <summary>
        /// If the axis is deleted
        /// </summary>
        public abstract bool Deleted 
        {
            get;
            set;
        }
        /// <summary>
        /// Position of the Lables
        /// </summary>
        public abstract eTickLabelPosition TickLabelPosition 
        {
            get;
            set;
        }
        /// <summary>
        /// The scaling value of the display units for the value axis
        /// </summary>
        public abstract double DisplayUnit
        {
            get;
            set;
        }
        /// <summary>
        /// Chart axis title
        /// </summary>
        internal protected ExcelChartTitle _title=null;
        /// <summary>
        /// Gives access to the charts title properties.
        /// </summary>
        public virtual ExcelChartTitle Title
        {
            get
            {                                
                return GetTitle();
            }
        }

        internal abstract ExcelChartTitle GetTitle();
        #region "Scaling"
        /// <summary>
        /// Minimum value for the axis.
        /// Null is automatic
        /// </summary>
        public abstract double? MinValue
        {
            get;
            set;
        }
        /// <summary>
        /// Max value for the axis.
        /// Null is automatic
        /// </summary>
        public abstract double? MaxValue
        {
            get;
            set;
        }
        /// <summary>
        /// Major unit for the axis.
        /// Null is automatic
        /// </summary>
        public abstract double? MajorUnit
        {
            get;
            set;
        }
        /// <summary>
        /// Major time unit for the axis.
        /// Null is automatic
        /// </summary>
        public abstract eTimeUnit? MajorTimeUnit
        {
            get;
            set;
        }
        /// <summary>
        /// Minor unit for the axis.
        /// Null is automatic
        /// </summary>
        public abstract double? MinorUnit
        {
            get;
            set;
        }
        /// <summary>
        /// Minor time unit for the axis.
        /// Null is automatic
        /// </summary>
        public abstract eTimeUnit? MinorTimeUnit
        {
            get;
            set;
        }
        /// <summary>
        /// The base for a logaritmic scale
        /// Null for a normal scale
        /// </summary>
        public abstract double? LogBase
        {
            get;
            set;
        }
        /// <summary>
        /// Axis orientation
        /// </summary>
        public abstract eAxisOrientation Orientation
        {
            get;
            set;
        }
        #endregion

        #region GridLines 
        ExcelDrawingBorder _majorGridlines = null; 
  
        /// <summary> 
        /// Major gridlines for the axis 
        /// </summary> 
        public ExcelDrawingBorder MajorGridlines
        { 
            get 
            { 
                if (_majorGridlines == null) 
                {  
                    _majorGridlines = new ExcelDrawingBorder(_chart, NameSpaceManager, TopNode,$"{_majorGridlinesPath}/{_nsPrefix}:spPr/a:ln", SchemaNodeOrder); 
                } 
                return _majorGridlines; 
            } 
        }
        ExcelDrawingEffectStyle _majorGridlineEffects = null;
        /// <summary> 
        /// Effects for major gridlines for the axis 
        /// </summary> 
        public ExcelDrawingEffectStyle MajorGridlineEffects
        {
            get
            {
                if (_majorGridlineEffects == null)
                {
                    _majorGridlineEffects = new ExcelDrawingEffectStyle(_chart, NameSpaceManager, TopNode, $"{_majorGridlinesPath}/{_nsPrefix}:spPr/a:effectLst", SchemaNodeOrder);
                }
                return _majorGridlineEffects;
            }
        }

        ExcelDrawingBorder _minorGridlines = null; 
  
        /// <summary> 
        /// Minor gridlines for the axis 
        /// </summary> 
        public ExcelDrawingBorder MinorGridlines
        { 
            get 
            { 
                if (_minorGridlines == null) 
                {  
                    _minorGridlines = new ExcelDrawingBorder(_chart, NameSpaceManager, TopNode,$"{_minorGridlinesPath}/{_nsPrefix}:spPr/a:ln", SchemaNodeOrder); 
                } 
                return _minorGridlines; 
            } 
        }
        ExcelDrawingEffectStyle _minorGridlineEffects = null;
        /// <summary> 
        /// Effects for minor gridlines for the axis 
        /// </summary> 
        public ExcelDrawingEffectStyle MinorGridlineEffects
        {
            get
            {
                if (_minorGridlineEffects == null)
                {
                    _minorGridlineEffects = new ExcelDrawingEffectStyle(_chart, NameSpaceManager, TopNode, $"{_minorGridlinesPath}/{_nsPrefix}:spPr/a:effectLst", SchemaNodeOrder);
                }
                return _minorGridlineEffects;
            }
        }
        /// <summary>
        /// True if the axis has major Gridlines
        /// </summary>
        public bool HasMajorGridlines
        {
            get
            {
                return ExistsNode(_majorGridlinesPath);
            }
        }
        /// <summary>
        /// True if the axis has minor Gridlines
        /// </summary>
        public bool HasMinorGridlines
        {
            get
            {
                return ExistsNode(_minorGridlinesPath);
            }
        }        
        /// <summary> 
        /// Removes Major and Minor gridlines from the Axis 
        /// </summary> 
        public void RemoveGridlines()
        { 
            RemoveGridlines(true,true); 
        }
        /// <summary>
        ///  Removes gridlines from the Axis
        /// </summary>
        /// <param name="removeMajor">Indicates if the Major gridlines should be removed</param>
        /// <param name="removeMinor">Indicates if the Minor gridlines should be removed</param>
        public void RemoveGridlines(bool removeMajor, bool removeMinor)
        { 
            if (removeMajor) 
            { 
                DeleteNode(_majorGridlinesPath); 
                _majorGridlines = null; 
            } 
  
            if (removeMinor) 
            { 
                DeleteNode(_minorGridlinesPath);    
                _minorGridlines = null; 
            } 
        }
        /// <summary>
        /// Adds gridlines and styles them according to the style selected in the StyleManager
        /// </summary>
        /// <param name="addMajor">Indicates if the Major gridlines should be added</param>
        /// <param name="addMinor">Indicates if the Minor gridlines should be added</param>
        public void AddGridlines(bool addMajor=true, bool addMinor=false)
        {
            if(addMajor)
            {
                CreateNode(_majorGridlinesPath);
                _chart.ApplyStyleOnPart(this, _chart._styleManager?.Style?.GridlineMajor);
            }
            if (addMinor)
            {
                CreateNode(_minorGridlinesPath);
                _chart.ApplyStyleOnPart(this, _chart._styleManager?.Style?.GridlineMinor);
            }
        }
        /// <summary>
        /// Adds the axis title and styles it according to the style selected in the StyleManager
        /// </summary>
        /// <param name="title"></param>
        public void AddTitle(string title)
        {
            Title.Text = title;
            _chart.ApplyStyleOnPart(Title, _chart._styleManager?.Style?.AxisTitle);
        }
        /// <summary>
        /// Removes the axis title
        /// </summary>
        public void RemoveTitle()
        {
            DeleteNode($"{_nsPrefix}:title");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        internal void ChangeAxisType(eAxisType type)
        {
            var children = XmlHelper.CopyToSchemaNodeOrder(ExcelChartAxisStandard._schemaNodeOrderDateShared, ExcelChartAxisStandard._schemaNodeOrderDate);
            RenameNode(TopNode, "c", "dateAx", children);            
        }
        #endregion
        internal XmlNode AddTitleNode()
        {
            var node = TopNode.SelectSingleNode($"{_nsPrefix}:title", NameSpaceManager);
            if (node == null)
            {
                node = CreateNode($"{_nsPrefix}:title");
                if (_chart._isChartEx == false)
                {
                    node.InnerXml = ExcelChartTitle.GetInitXml(_nsPrefix);
                }
            }
            return node;
        }
        void IStyleMandatoryProperties.SetMandatoryProperties()
        {
            TextBody.Anchor = eTextAnchoringType.Center;
            TextBody.AnchorCenter = true;
            TextBody.WrapText = eTextWrappingType.Square;
            TextBody.VerticalTextOverflow = eTextVerticalOverflow.Ellipsis;
            TextBody.ParagraphSpacing = true;
            TextBody.Rotation = 0;

            if (Font.Kerning == 0) Font.Kerning = 12;
            Font.Bold = Font.Bold; //Must be set

            CreatespPrNode($"{_nsPrefix}:spPr");
        }
    }
}
