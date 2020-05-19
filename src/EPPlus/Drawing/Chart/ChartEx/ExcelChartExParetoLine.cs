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
using OfficeOpenXml.Drawing.Style.Effect;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelChartExParetoLine : ExcelDrawingBorder
    {
        private readonly ExcelChart _chart;
        internal ExcelChartExParetoLine(ExcelChart chart, XmlNamespaceManager nsm, XmlNode node) : base(chart, nsm, node, "cx:spPr/a:ln", new string[] { "spPr", "axisId" })
        {
            _chart = chart;
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
                    _effect = new ExcelDrawingEffectStyle(_chart, NameSpaceManager, TopNode, "cx:spPr/a:effectLst", SchemaNodeOrder);
                }
                return _effect;
            }
        }
    }
}