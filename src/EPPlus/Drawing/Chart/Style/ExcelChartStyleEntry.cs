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
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Effect;
using System.Xml;
using OfficeOpenXml.Utils.Extensions;
using System;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Drawing.Chart.Style
{
    /// <summary>
    /// A style entry for a chart part.
    /// </summary>
    public class ExcelChartStyleEntry : XmlHelper
    {
        string _fillReferencePath = "{0}/cs:fillRef";
        string _borderReferencePath = "{0}/cs:lnRef ";
        string _effectReferencePath = "{0}/cs:effectRef";
        string _fontReferencePath = "{0}/cs:fontRef";

        string _richTextPath = "{0}/cs:rich";
        string _fillPath = "{0}/cs:spPr";
        string _borderPath = "{0}/cs:spPr/a:ln";
        string _effectPath = "{0}/cs:spPr/a:effectLst";
        string _scene3DPath = "{0}/cs:spPr/a:scene3d";
        string _sp3DPath = "{0}/cs:spPr/a:sp3d";

        string _defaultTextRunPath = "{0}/cs:defRPr";
        string _defaultTextBodyPath = "{0}/cs:bodyPr";
        private readonly IPictureRelationDocument _pictureRelationDocument;
        internal ExcelChartStyleEntry(XmlNamespaceManager nsm, XmlNode topNode, string path, IPictureRelationDocument pictureRelationDocument) : base(nsm, topNode)
        {
            SchemaNodeOrder = new string[] { "lnRef", "fillRef", "effectRef", "fontRef", "spPr", "noFill", "solidFill", "blipFill", "gradFill", "noFill", "pattFill","ln", "defRPr" };
            _fillReferencePath = string.Format(_fillReferencePath, path);
            _borderReferencePath = string.Format(_borderReferencePath, path);
            _effectReferencePath = string.Format(_effectReferencePath, path);
            _fontReferencePath = string.Format(_fontReferencePath, path);

            _richTextPath = string.Format(_richTextPath, path);
            _fillPath = string.Format(_fillPath, path);
            _borderPath = string.Format(_borderPath, path);
            _effectPath = string.Format(_effectPath, path);
            _scene3DPath = string.Format(_scene3DPath, path);
            _sp3DPath = string.Format(_sp3DPath, path);

            _defaultTextRunPath = string.Format(_defaultTextRunPath, path);
            _defaultTextBodyPath = string.Format(_defaultTextBodyPath, path);
            _pictureRelationDocument = pictureRelationDocument;
        }
        private ExcelChartStyleReference _borderReference = null;
        /// Border reference. 
        /// Contains an index reference to the theme and a color to be used in border styling
        public ExcelChartStyleReference BorderReference
        {
            get
            {
                if (_borderReference == null)
                {
                    _borderReference = new ExcelChartStyleReference(NameSpaceManager, TopNode, _borderReferencePath);
                }
                return _borderReference;
            }
        }
        private ExcelChartStyleReference _fillReference = null;
        /// <summary>
        /// Fill reference. 
        /// Contains an index reference to the theme and a fill color to be used in fills
        /// </summary>
        public ExcelChartStyleReference FillReference
        {
            get
            {
                if (_fillReference == null)
                {
                    _fillReference = new ExcelChartStyleReference(NameSpaceManager, TopNode, _fillReferencePath);
                }
                return _fillReference;
            }
        }
        private ExcelChartStyleReference _effectReference = null;
        /// <summary>
        /// Effect reference. 
        /// Contains an index reference to the theme and a color to be used in effects
        /// </summary>
        public ExcelChartStyleReference EffectReference
        {
            get
            {
                if (_effectReference == null)
                {
                    _effectReference = new ExcelChartStyleReference(NameSpaceManager, TopNode, _effectReferencePath);
                }
                return _effectReference;
            }
        }
        ExcelChartStyleFontReference _fontReference = null;
        /// <summary>
        /// Font reference. 
        /// Contains an index reference to the theme and a color to be used for font styling
        /// </summary>
        public ExcelChartStyleFontReference FontReference
        {
            get
            {
                if (_fontReference == null)
                {
                    _fontReference = new ExcelChartStyleFontReference(NameSpaceManager, TopNode, _fontReferencePath);
                }
                return _fontReference;
            }
        }

        private ExcelDrawingFill _fill = null;
        /// <summary>
        /// Reference to fill settings for a chart part
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if(_fill==null)
                {
                    _fill = new ExcelDrawingFill(_pictureRelationDocument, NameSpaceManager, TopNode, _fillPath, SchemaNodeOrder);
                }
                return _fill;
            }
        }
        private ExcelDrawingBorder _border = null;
        /// <summary>
        /// Reference to border settings for a chart part
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(_pictureRelationDocument, NameSpaceManager, TopNode, _borderPath, SchemaNodeOrder);
                }
                return _border;
            }
        }
        private ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// Reference to border settings for a chart part
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (_effect == null)
                {
                    _effect = new ExcelDrawingEffectStyle(_pictureRelationDocument, NameSpaceManager, TopNode, _effectPath, SchemaNodeOrder);
                }
                return _effect;
            }
        }
        private ExcelDrawing3D _threeD = null;
        /// <summary>
        /// Reference to 3D effect settings for a chart part
        /// </summary>
        public ExcelDrawing3D ThreeD
        {
            get
            {
                if (_threeD == null)
                {
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, _fillPath, SchemaNodeOrder);
                }
                return _threeD;
            }
        }
        private ExcelTextRun _defaultTextRun = null;
        /// <summary>
        /// Reference to default text run settings for a chart part
        /// </summary>
        public ExcelTextRun DefaultTextRun
        {
            get
            {
                if (_defaultTextRun == null)
                {
                    _defaultTextRun = new ExcelTextRun(NameSpaceManager, TopNode, _defaultTextRunPath);
                }
                return _defaultTextRun;
                
            }
        }
        private ExcelTextBody _defaultTextBody = null;
        /// <summary>
        /// Reference to default text body run settings for a chart part
        /// </summary>
        public ExcelTextBody DefaultTextBody
        {
            get 
            {
                if (_defaultTextBody == null)
                {
                    _defaultTextBody = new ExcelTextBody(NameSpaceManager, TopNode, _defaultTextBodyPath);
                }
                return _defaultTextBody;

            }
        }
        /// <summary>
        /// Modifier for the chart
        /// </summary>
        public eStyleEntryModifier Modifier
        {
            get
            {
                var split = GetXmlNodeString("@mods").Split(' ');
                eStyleEntryModifier ret=0;
                foreach(var v in split)
                {
                    ret |= v.ToEnum<eStyleEntryModifier>(0);
                }
                return ret;
            }
            set
            {
                string s = "";
                foreach(eStyleEntryModifier e in Enum.GetValues(typeof(eStyleEntryModifier)))
                {
                    if ((int)(value & e) != 0)
                    {
                        s += e.ToString() + " ";
                    }
                }
                if(s=="")
                {
                    ((XmlElement)TopNode).RemoveAttribute("mods"); 
                }
                else
                {
                    SetXmlNodeString("@mods", s.Substring(0,s.Length-1));
                }
            }
        }
        /// <summary>
        /// True if the entry has fill styles
        /// </summary>
        public bool HasFill
        {
            get
            {
                return ExistsNode($"{_fillPath}/a:noFill") ||
                       ExistsNode($"{_fillPath}/a:solidFill") ||
                       ExistsNode($"{_fillPath}/a:gradFill") ||
                       ExistsNode($"{_fillPath}/a:pattFill") ||
                       ExistsNode($"{_fillPath}/a:blipFill");
            }
        }
        /// <summary>
        /// True if the entry has border styles
        /// </summary>
        public bool HasBorder
        {
            get
            {
                return ExistsNode(_borderPath);
            }
        }
        /// <summary>
        /// True if the entry effects styles
        /// </summary>
        public bool HasEffect
        {
            get
            {
                return ExistsNode(_effectPath);
            }
        }
        /// <summary>
        /// True if the entry has 3D styles
        /// </summary>
        public bool HasThreeD
        {
            get
            {
                return ExistsNode(_scene3DPath) || ExistsNode(_sp3DPath);
            }
        }

        /// <summary>
        /// True if the entry has text body styles
        /// </summary>
        public bool HasTextBody
        {
            get
            {
                return ExistsNode(_defaultTextBodyPath);
            }
        }
        /// <summary>
        /// True if the entry has richtext
        /// </summary>
        public bool HasRichText
        {
            get
            {
                return ExistsNode(_richTextPath);
            }
        }

        /// <summary>
        /// True if the entry has text run styles
        /// </summary>
        public bool HasTextRun
        {
            get
            {
                return ExistsNode(_defaultTextRunPath);
            }
        }
    }
}