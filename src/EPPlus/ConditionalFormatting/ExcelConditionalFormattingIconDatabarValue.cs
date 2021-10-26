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
using System.Linq;
using System.Text;
using System.Drawing;
using System.Xml;
using OfficeOpenXml.Utils;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Security;

namespace OfficeOpenXml.ConditionalFormatting
{
	/// <summary>
	/// 18.3.1.11 cfvo (Conditional Format Value Object)
	/// Describes the values of the interpolation points in a gradient scale.
	/// </summary>
	public class ExcelConditionalFormattingIconDataBarValue
		: XmlHelper
	{
		/****************************************************************************************/

		#region Private Properties
		private eExcelConditionalFormattingRuleType _ruleType;
		private ExcelWorksheet _worksheet;
		#endregion Private Properties

		/****************************************************************************************/

		#region Constructors
    /// <summary>
    /// Initialize the cfvo (§18.3.1.11) node
    /// </summary>
    /// <param name="type"></param>
    /// <param name="value"></param>
    /// <param name="formula"></param>
    /// <param name="ruleType"></param>
    /// <param name="address"></param>
    /// <param name="priority"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode">The cfvo (§18.3.1.11) node parent. Can be any of the following:
    /// colorScale (§18.3.1.16); dataBar (§18.3.1.28); iconSet (§18.3.1.49)</param>
    /// <param name="namespaceManager"></param>
		internal ExcelConditionalFormattingIconDataBarValue(
			eExcelConditionalFormattingValueObjectType type,
			double value,
			string formula,
			eExcelConditionalFormattingRuleType ruleType,
            ExcelAddress address,
            int priority,
			ExcelWorksheet worksheet,
			XmlNode itemElementNode,
			XmlNamespaceManager namespaceManager)
			: this(
            ruleType,
            address,
            worksheet,
            itemElementNode,
			namespaceManager)
		{
            // Check if the parent does not exists
			if (itemElementNode == null)
			{
				// Get the parent node path by the rule type
				string parentNodePath = ExcelConditionalFormattingValueObjectType.GetParentPathByRuleType(
					ruleType);

				// Check for en error (rule type does not have <cfvo>)
				if (parentNodePath == string.Empty)
				{
					throw new Exception(
						ExcelConditionalFormattingConstants.Errors.MissingCfvoParentNode);
				}

				// Point to the <cfvo> parent node
        itemElementNode = _worksheet.WorksheetXml.SelectSingleNode(
					string.Format(
						"//{0}[{1}='{2}']/{3}[{4}='{5}']/{6}",
					// {0}
						ExcelConditionalFormattingConstants.Paths.ConditionalFormatting,
					// {1}
						ExcelConditionalFormattingConstants.Paths.SqrefAttribute,
					// {2}
						address.Address,
					// {3}
						ExcelConditionalFormattingConstants.Paths.CfRule,
					// {4}
						ExcelConditionalFormattingConstants.Paths.PriorityAttribute,
					// {5}
						priority,
					// {6}
						parentNodePath),
					_worksheet.NameSpaceManager);

				// Check for en error (rule type does not have <cfvo>)
                if (itemElementNode == null)
				{
					throw new Exception(
						ExcelConditionalFormattingConstants.Errors.MissingCfvoParentNode);
				}
			}

            TopNode = itemElementNode;

			// Save the attributes
			RuleType = ruleType;
			Type = type;
			Value = value;
			Formula = formula;
		}
    /// <summary>
    /// Initialize the cfvo (§18.3.1.11) node
    /// </summary>
    /// <param name="ruleType"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode">The cfvo (§18.3.1.11) node parent. Can be any of the following:
    /// colorScale (§18.3.1.16); dataBar (§18.3.1.28); iconSet (§18.3.1.49)</param>
    /// <param name="namespaceManager"></param>
        internal ExcelConditionalFormattingIconDataBarValue(
            eExcelConditionalFormattingRuleType ruleType,
            ExcelAddress address,
            ExcelWorksheet worksheet,
            XmlNode itemElementNode,
            XmlNamespaceManager namespaceManager)
            : base(
                namespaceManager,
                itemElementNode)
        {
            Require.Argument(address).IsNotNull("address");
            Require.Argument(worksheet).IsNotNull("worksheet");

            // Save the worksheet for private methods to use
            _worksheet = worksheet;

            // Schema order list
            SchemaNodeOrder = new string[]
			{
                ExcelConditionalFormattingConstants.Nodes.Cfvo,
			};

            //Check if the parent does not exists
            if (itemElementNode == null)
            {
                // Get the parent node path by the rule type
                string parentNodePath = ExcelConditionalFormattingValueObjectType.GetParentPathByRuleType(
                    ruleType);

                // Check for en error (rule type does not have <cfvo>)
                if (parentNodePath == string.Empty)
                {
                    throw new Exception(
                        ExcelConditionalFormattingConstants.Errors.MissingCfvoParentNode);
                }
            }
            RuleType = ruleType;            
        }
		/// <summary>
		/// Initialize the <see cref="ExcelConditionalFormattingColorScaleValue"/>
		/// </summary>
		/// <param name="type"></param>
		/// <param name="value"></param>
		/// <param name="formula"></param>
		/// <param name="ruleType"></param>
		/// <param name="priority"></param>
		/// <param name="address"></param>
		/// <param name="worksheet"></param>
		/// <param name="namespaceManager"></param>
		internal ExcelConditionalFormattingIconDataBarValue(
			eExcelConditionalFormattingValueObjectType type,
			double value,
			string formula,
			eExcelConditionalFormattingRuleType ruleType,
            ExcelAddress address,
            int priority,
			ExcelWorksheet worksheet,
			XmlNamespaceManager namespaceManager)
			: this(
				type,
				value,
				formula,
				ruleType,
                address,
                priority,
				worksheet,
				null,
				namespaceManager)
		{
            
		}
		/// <summary>
		/// Initialize the <see cref="ExcelConditionalFormattingColorScaleValue"/>
		/// </summary>
		/// <param name="type"></param>
		/// <param name="color"></param>
		/// <param name="ruleType"></param>
		/// <param name="priority"></param>
		/// <param name="address"></param>
		/// <param name="worksheet"></param>
		/// <param name="namespaceManager"></param>
		internal ExcelConditionalFormattingIconDataBarValue(
			eExcelConditionalFormattingValueObjectType type,
			Color color,
			eExcelConditionalFormattingRuleType ruleType,
            ExcelAddress address,
            int priority,
			ExcelWorksheet worksheet,
			XmlNamespaceManager namespaceManager)
			: this(
				type,
				0,
				null,
				ruleType,
                address,
                priority,
				worksheet,
				null,
				namespaceManager)
		{
		}
		#endregion Constructors

		/****************************************************************************************/

		#region Methods
        #endregion

        /****************************************************************************************/

		#region Exposed Properties
		/// <summary>
		/// Rule type
		/// </summary>
		internal eExcelConditionalFormattingRuleType RuleType
		{
			get { return _ruleType; }
			set { _ruleType = value; }
		}

		/// <summary>
		/// Value type
		/// </summary>
		public eExcelConditionalFormattingValueObjectType Type
		{
			get
			{
				var typeAttribute = GetXmlNodeString(ExcelConditionalFormattingConstants.Paths.TypeAttribute);

				return ExcelConditionalFormattingValueObjectType.GetTypeByAttrbiute(typeAttribute);
			}
			set
			{
                if ((_ruleType==eExcelConditionalFormattingRuleType.ThreeIconSet || _ruleType==eExcelConditionalFormattingRuleType.FourIconSet || _ruleType==eExcelConditionalFormattingRuleType.FiveIconSet) &&
                    (value == eExcelConditionalFormattingValueObjectType.Min || value == eExcelConditionalFormattingValueObjectType.Max))
                {
                    throw(new ArgumentException("Value type can't be Min or Max for icon sets"));
                }
                SetXmlNodeString(ExcelConditionalFormattingConstants.Paths.TypeAttribute, value.ToString().ToLower(CultureInfo.InvariantCulture));                
			}
		}

        /// <summary>
        /// Greater Than Or Equal 
        /// </summary>
        public bool GreaterThanOrEqualTo
        {
            get
            {
                return GetXmlNodeBool(ExcelConditionalFormattingConstants.Paths.GteAttribute);
            }

            set
            {
                SetXmlNodeString(  
                    ExcelConditionalFormattingConstants.Paths.GteAttribute,
                    (value == false) ? "0" : string.Empty,
                    true);
            }
        }



        /// <summary>
        /// The value
        /// </summary>
        public Double Value
		{
			get
			{
                if ((Type == eExcelConditionalFormattingValueObjectType.Num)
                    || (Type == eExcelConditionalFormattingValueObjectType.Percent)
                    || (Type == eExcelConditionalFormattingValueObjectType.Percentile))
                {
                    return GetXmlNodeDouble(ExcelConditionalFormattingConstants.Paths.ValAttribute);
                }
                else
                {
                    return 0;
                }
            }
			set
			{
				string valueToStore = string.Empty;

				// Only some types use the @val attribute
				if ((Type == eExcelConditionalFormattingValueObjectType.Num)
					|| (Type == eExcelConditionalFormattingValueObjectType.Percent)
					|| (Type == eExcelConditionalFormattingValueObjectType.Percentile))
				{
					valueToStore = value.ToString(CultureInfo.InvariantCulture);
				}

                SetXmlNodeString(ExcelConditionalFormattingConstants.Paths.ValAttribute, valueToStore);
			}
		}

		/// <summary>
		/// The Formula of the Object Value (uses the same attribute as the Value)
		/// </summary>
		public string Formula
		{
			get
			{
				// Return empty if the Object Value type is not Formula
				if (Type != eExcelConditionalFormattingValueObjectType.Formula)
				{
					return string.Empty;
				}

				// Excel stores the formula in the @val attribute
				return GetXmlNodeString(ExcelConditionalFormattingConstants.Paths.ValAttribute);
			}
			set
			{
				// Only store the formula if the Object Value type is Formula
				if (Type == eExcelConditionalFormattingValueObjectType.Formula)
				{
                    SetXmlNodeString(ExcelConditionalFormattingConstants.Paths.ValAttribute, value);
				}
			}
		}
		#endregion Exposed Properties

		/****************************************************************************************/
	}
}