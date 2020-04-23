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
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelChartExDataPoint : XmlHelper, IDrawingStyleBase
    {
        ExcelChart _chart;
        internal ExcelChartExDataPoint(ExcelChart chart, XmlNamespaceManager ns, XmlNode topNode) : base(ns, topNode)
        {
            SchemaNodeOrder=new string[]{"spPr"};
            Index = GetXmlNodeInt(indexPath);
        }
        internal ExcelChartExDataPoint(ExcelChart chart, XmlNamespaceManager ns, XmlNode topNode, int index) : base(ns, topNode)
        {
            SchemaNodeOrder = new string[] { "spPr" };
            Index = index;
        }

        internal const string topNodePath = "cx:dataPt";
        const string indexPath = "@idx";
        /// <summary>
        /// The index of the datapoint
        /// </summary>
        public int Index
        {
            get;
            private set;
        }

        ExcelDrawingFill _fill = null;
        /// <summary>
        /// A reference to fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(_chart, NameSpaceManager, TopNode, "cx:spPr", SchemaNodeOrder);
                }
                return _fill;
            }
        }
        ExcelDrawingBorder _line = null;
        /// <summary>
        /// A reference to line properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_line == null)
                {
                    _line = new ExcelDrawingBorder(_chart, NameSpaceManager, TopNode, "cx:spPr/a:ln", SchemaNodeOrder);
                }
                return _line;
            }
        }
        private ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// A reference to line properties
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (_effect == null)
                {
                    _effect = new ExcelDrawingEffectStyle(_chart, NameSpaceManager, TopNode, "cx:spPr/a:effectLst", SchemaNodeOrder);
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
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, "cx:spPr", SchemaNodeOrder);
                }
                return _threeD;
            }
        }
        void IDrawingStyleBase.CreatespPr()
        {
            CreatespPrNode();
        }

    }
}