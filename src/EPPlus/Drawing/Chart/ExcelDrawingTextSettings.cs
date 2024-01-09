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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using System.Xml;
using OfficeOpenXml.Drawing.Style.Effect;
namespace OfficeOpenXml.Drawing.Chart
{
	/// <summary>
	/// Text settings for drawing objects.
	/// </summary>
	public class ExcelDrawingTextSettings : XmlHelper
	{
		private ExcelChart _chart;
		private string _topPath;

		internal ExcelDrawingTextSettings(ExcelChart chart, XmlNamespaceManager ns, XmlNode topNode, string topPath, string[] schemaNodeOrder) : base(ns, topNode)
		{
			_chart = chart;
			_topPath = topPath;
			SchemaNodeOrder = schemaNodeOrder;
		}
		ExcelDrawingFill _fill = null;
		/// <summary>
		/// The Fill style
		/// </summary>
		public ExcelDrawingFill Fill
		{
			get
			{
				if (_fill == null)
				{
					_fill = new ExcelDrawingFill(_chart, NameSpaceManager, TopNode, _topPath, SchemaNodeOrder);
				}
				return _fill;
			}
		}
		ExcelDrawingBorder _border = null;
		/// <summary>
		/// The text outline style.
		/// </summary>
		public ExcelDrawingBorder Outline
		{
			get
			{
				if (_border == null)
				{
					_border = new ExcelDrawingBorder(_chart, NameSpaceManager, TopNode, $"{_topPath}/a:ln", SchemaNodeOrder);
				}
				return _border;
			}
		}

		ExcelDrawingEffectStyle _effect = null;
		/// <summary>
		/// Text effects
		/// </summary>
		public ExcelDrawingEffectStyle Effect
		{
			get
			{
				if (_effect == null)
				{
					_effect = new ExcelDrawingEffectStyle(_chart, NameSpaceManager, TopNode, $"{_topPath}/a:effectLst", SchemaNodeOrder);
				}
				return _effect;
			}
		}
	}
}