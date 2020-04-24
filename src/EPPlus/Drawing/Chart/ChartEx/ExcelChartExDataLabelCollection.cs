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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelChartExDataLabelCollection : ExcelChartExDataLabel, IDrawingStyle, IEnumerable<ExcelChartExDataLabelItem>
    {
        internal ExcelChartExDataLabelCollection(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, string[] schemaNodeOrder) : 
            base(chart, ns, node)
        {
            _chart = chart;
            AddSchemaNodeOrder(schemaNodeOrder, new string[]{ "numFmt","spPr", "txPr", "visibility", "separator"});
        }
        const string _seriesNameVisiblePath = "cx:visibility/@seriesName";
        public bool SeriesNameVisible
        { 
            get
            {
                return GetXmlNodeBool(_seriesNameVisiblePath);
            }
            set
            {
                SetXmlNodeBool(_seriesNameVisiblePath, value);
            }
        }
        const string _categoryNameVisiblePath = "cx:visibility/@categoryName";
        public bool CategoryNameVisible
        {
            get
            {
                return GetXmlNodeBool(_categoryNameVisiblePath);
            }
            set
            {
                SetXmlNodeBool(_categoryNameVisiblePath, value);
            }
        }
        const string _valueVisiblePath = "cx:visibility/@value";
        public bool ValueVisible
        {
            get
            {
                return GetXmlNodeBool(_valueVisiblePath);
            }
            set
            {
                SetXmlNodeBool(_valueVisiblePath, value);
            }
        }
        const string _separatorPath = "cx:separator";
        public string Separator 
        {
            get
            {
                return GetXmlNodeString(_separatorPath);
            }
            set
            {
                SetXmlNodeString(_separatorPath, value, true);
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
        private void CreateDefaultText()
        {
            if (TopNode.SelectSingleNode("cx:txPr", NameSpaceManager) == null)
            {
                if (!ExistNode("cx:spPr"))
                {
                    var spNode = CreateNode("cx:spPr");
                    spNode.InnerXml = "<a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/>";
                }
                var node = CreateNode("cx:txPr");
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

        void IDrawingStyleBase.CreatespPr()
        {
            CreatespPrNode("cx:spPr");
        }

        public IEnumerator<ExcelChartExDataLabelItem> GetEnumerator()
        {
            throw new System.NotImplementedException();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new System.NotImplementedException();
        }
    }
}
