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
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style.Dxf;

namespace OfficeOpenXml.ConditionalFormatting
{
  /// <summary>
  /// 
  /// </summary>
  public abstract class ExcelConditionalFormattingRule
    : XmlHelper,
    IExcelConditionalFormattingRule
  {
    /****************************************************************************************/

    #region Private Properties
    private eExcelConditionalFormattingRuleType? _type;
    private ExcelWorksheet _worksheet;

    /// <summary>
    /// Sinalize that we are in a Cnaging Priorities opeartion so that we won't enter
    /// a recursive loop.
    /// </summary>
    private static bool _changingPriority = false;
    #endregion Private Properties

    /****************************************************************************************/

    #region Constructors
    /// <summary>
    /// Initialize the <see cref="ExcelConditionalFormattingRule"/>
    /// </summary>
    /// <param name="type"></param>
    /// <param name="address"></param>
    /// <param name="priority">Used also as the cfRule unique key</param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode"></param>
    /// <param name="namespaceManager"></param>
    internal ExcelConditionalFormattingRule(
      eExcelConditionalFormattingRuleType type,
      ExcelAddress address,
      int priority,
      ExcelWorksheet worksheet,
      XmlNode itemElementNode,
      XmlNamespaceManager namespaceManager)
      : base(
        namespaceManager,
        itemElementNode)
    {
      Require.Argument(address).IsNotNull("address");
      Require.Argument(worksheet).IsNotNull("worksheet");

      _type = type;
      _worksheet = worksheet;
      SchemaNodeOrder = _worksheet.SchemaNodeOrder;

      if (itemElementNode == null)
      {
        // Create/Get the <cfRule> inside <conditionalFormatting>
        itemElementNode = CreateComplexNode(
          _worksheet.WorksheetXml.DocumentElement,
          string.Format(
            "{0}[{1}='{2}']/{1}='{2}'/{3}[{4}='{5}']/{4}='{5}'",
          //{0}
            ExcelConditionalFormattingConstants.Paths.ConditionalFormatting,
          // {1}
            ExcelConditionalFormattingConstants.Paths.SqrefAttribute,
          // {2}
            address.AddressSpaceSeparated,          //CF node don't what to have comma between multi addresses, use space instead.
          // {3}
            ExcelConditionalFormattingConstants.Paths.CfRule,
          //{4}
            ExcelConditionalFormattingConstants.Paths.PriorityAttribute,
          //{5}
            priority));
      }

      // Point to <cfRule>
      TopNode = itemElementNode;

      Address = address;
      Priority = priority;
      Type = type;  
      if (DxfId >= 0 && DxfId < worksheet.Workbook.Styles.Dxfs.Count)
      {
          worksheet.Workbook.Styles.Dxfs[DxfId].AllowChange = true;  //This Id is referenced by CF, so we can use it when we save.
          _style = ((ExcelDxfStyleBase)worksheet.Workbook.Styles.Dxfs[DxfId]).ToDxfConditionalFormattingStyle();    //Clone, so it can be altered without affecting other dxf styles
      }
    }

    /// <summary>
    /// Initialize the <see cref="ExcelConditionalFormattingRule"/>
    /// </summary>
    /// <param name="type"></param>
    /// <param name="address"></param>
    /// <param name="priority"></param>
    /// <param name="worksheet"></param>
    /// <param name="namespaceManager"></param>
    internal ExcelConditionalFormattingRule(
      eExcelConditionalFormattingRuleType type,
      ExcelAddress address,
      int priority,
      ExcelWorksheet worksheet,
      XmlNamespaceManager namespaceManager)
      : this(
        type,
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
    #endregion Methods

    /****************************************************************************************/

    #region Exposed Properties
    /// <summary>
    /// Get the &lt;cfRule&gt; node
    /// </summary>
    public XmlNode Node
    {
      get { return TopNode; }
    }

    /// <summary>
    /// The address of the conditional formatting rule
    /// </summary>
    /// <remarks>
    /// The address is stored in a parent node called &lt;conditionalFormatting&gt; in the
    /// @sqref attribute. Excel groups rules that have the same address inside one node.
    /// </remarks>
    public ExcelAddress Address
    {
      get
      {
        return new ExcelAddress(
          Node.ParentNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Sqref].Value.Replace(" ",",")
          );
      }
      set
      {
        // Check if the address is to be changed
        if (!Address.Equals(value))
        {
          // Save the old parente node
          XmlNode oldParentNode = Node.ParentNode;

          // Create/Get the new <conditionalFormatting> parent node
          XmlNode newParentNode = CreateComplexNode(
            _worksheet.WorksheetXml.DocumentElement,
            string.Format(
              "{0}[{1}='{2}']/{1}='{2}'",
            //{0}
              ExcelConditionalFormattingConstants.Paths.ConditionalFormatting,
            // {1}
              ExcelConditionalFormattingConstants.Paths.SqrefAttribute,
            // {2}
              value.AddressSpaceSeparated)); 

          // Move the <cfRule> node to the new <conditionalFormatting> parent node
          TopNode = newParentNode.AppendChild(Node);

          // Check if the old <conditionalFormatting> parent node has <cfRule> node inside it
          if (!oldParentNode.HasChildNodes)
          {
            // Remove the old parent node
            oldParentNode.ParentNode.RemoveChild(oldParentNode);
          }
        }
      }
    }
        /// <summary>
        /// Indicates that the conditional formatting is associated with a PivotTable
        /// </summary>
        public bool PivotTable
        {
            get
            {
                return GetXmlNodeBool("../@pivot");
            }
            set
            {
                SetXmlNodeBool("../@pivot", value, false);
            }
        }
    /// <summary>
    /// Type of conditional formatting rule.
    /// </summary>
    public eExcelConditionalFormattingRuleType Type
    {
      get
      {
        // Transform the @type attribute to EPPlus Rule Type (slighty diferente)
        if(_type==null)
        {
          _type = ExcelConditionalFormattingRuleType.GetTypeByAttrbiute(
          GetXmlNodeString(ExcelConditionalFormattingConstants.Paths.TypeAttribute),
          TopNode,
          _worksheet.NameSpaceManager);
        }
        return (eExcelConditionalFormattingRuleType)_type;
      }
      internal set
      {
          _type = value;
          // Transform the EPPlus Rule Type to @type attribute (slighty diferente)
        SetXmlNodeString(
          ExcelConditionalFormattingConstants.Paths.TypeAttribute,
          ExcelConditionalFormattingRuleType.GetAttributeByType(value),
          true);
      }
    }

    /// <summary>
    /// The priority of the rule. 
    /// A lower values are higher priority than higher values, where 1 is the highest priority.
    /// </summary>
    public int Priority
    {
      get
      {
        return GetXmlNodeInt(
          ExcelConditionalFormattingConstants.Paths.PriorityAttribute);
      }
      set
      {
        // Save the current CF rule priority
        int priority = Priority;

        // Check if the @priority is to be changed
        if (priority != value)
        {
          // Check if we are not already inside a "Change Priority" operation
          if (!_changingPriority)
          {
            if (value < 1)
            {
              throw new IndexOutOfRangeException(
                ExcelConditionalFormattingConstants.Errors.InvalidPriority);
            }

            // Sinalize that we are already changing cfRules priorities
            _changingPriority = true;

            // Check if we lowered the priority
            if (priority > value)
            {
              for (int i = priority - 1; i >= value; i--)
              {
                var cfRule = _worksheet.ConditionalFormatting.RulesByPriority(i);

                if (cfRule != null)
                {
                  cfRule.Priority++;
                }
              }
            }
            else
            {
              for (int i = priority + 1; i <= value; i++)
              {
                var cfRule = _worksheet.ConditionalFormatting.RulesByPriority(i);

                if (cfRule != null)
                {
                  cfRule.Priority--;
                }
              }
            }

            // Sinalize that we are no longer changing cfRules priorities
            _changingPriority = false;
          }

          // Change the priority in the XML
          SetXmlNodeString(
            ExcelConditionalFormattingConstants.Paths.PriorityAttribute,
            value.ToString(),
            true);
        }
      }
    }

        /// <summary>
        /// If this property is true, no rules with lower priority shall be applied over this rule,
        /// when this rule evaluates to true.
        /// </summary>
        public bool StopIfTrue
    {
      get
      {
        return GetXmlNodeBool(
          ExcelConditionalFormattingConstants.Paths.StopIfTrueAttribute);
      }
      set
      {
        SetXmlNodeString(
          ExcelConditionalFormattingConstants.Paths.StopIfTrueAttribute,
          (value == true) ? "1" : string.Empty,
          true);
      }
    }

        /// <summary>
        /// The DxfId (Differential Formatting style id)
        /// </summary>
        internal int DxfId
    {
      get
      {
        return GetXmlNodeInt(
          ExcelConditionalFormattingConstants.Paths.DxfIdAttribute);
      }
      set
      {
        SetXmlNodeString(
          ExcelConditionalFormattingConstants.Paths.DxfIdAttribute,
          (value == int.MinValue) ? string.Empty : value.ToString(),
          true);
      }
    }
    internal ExcelDxfStyleConditionalFormatting _style = null;
    /// <summary>
    /// The style
    /// </summary>
    public ExcelDxfStyleConditionalFormatting Style
    {
        get
        {
            if (_style == null)
            {
                _style = new ExcelDxfStyleConditionalFormatting(NameSpaceManager, null, _worksheet.Workbook.Styles, null);                
            }
            return _style;
        }
    }        
    /// <summary>
    /// StdDev (zero is not allowed and will be converted to 1)
    /// </summary>
    public UInt16 StdDev
    {
      get
      {
        return Convert.ToUInt16(GetXmlNodeInt(
          ExcelConditionalFormattingConstants.Paths.StdDevAttribute));
      }
      set
      {
        SetXmlNodeString(
          ExcelConditionalFormattingConstants.Paths.StdDevAttribute,
          (value == 0) ? "1" : value.ToString(),
          true);
      }
    }

    /// <summary>
    /// Rank (zero is not allowed and will be converted to 1)
    /// </summary>
    public UInt16 Rank
    {
      get
      {
        return Convert.ToUInt16(GetXmlNodeInt(
          ExcelConditionalFormattingConstants.Paths.RankAttribute));
      }
      set
      {
        SetXmlNodeString(
          ExcelConditionalFormattingConstants.Paths.RankAttribute,
          (value == 0) ? "1" : value.ToString(),
          true);
      }
    }
    #endregion Exposed Properties

    /****************************************************************************************/

    #region Internal Properties
    /// <summary>
    /// Above average
    /// </summary>
    internal protected bool? AboveAverage
    {
      get
      {
        bool? aboveAverage = GetXmlNodeBoolNullable(
          ExcelConditionalFormattingConstants.Paths.AboveAverageAttribute);

        // Above Avarege if TRUE or if attribute does not exists
        return (aboveAverage == true) || (aboveAverage == null);
      }
      set
      {
        string aboveAverageValue = string.Empty;

        // Only the types that needs the @AboveAverage
        if ((_type == eExcelConditionalFormattingRuleType.BelowAverage)
          || (_type == eExcelConditionalFormattingRuleType.BelowOrEqualAverage)
          || (_type == eExcelConditionalFormattingRuleType.BelowStdDev))
        {
          aboveAverageValue = "0";
        }

        SetXmlNodeString(
          ExcelConditionalFormattingConstants.Paths.AboveAverageAttribute,
          aboveAverageValue,
          true);
      }
    }

    /// <summary>
    /// EqualAverage
    /// </summary>
    internal protected bool? EqualAverage
    {
      get
      {
        bool? equalAverage = GetXmlNodeBoolNullable(
          ExcelConditionalFormattingConstants.Paths.EqualAverageAttribute);

        // Equal Avarege only if TRUE
        return (equalAverage == true);
      }
      set
      {
        string equalAverageValue = string.Empty;

        // Only the types that needs the @EqualAverage
        if ((_type == eExcelConditionalFormattingRuleType.AboveOrEqualAverage)
          || (_type == eExcelConditionalFormattingRuleType.BelowOrEqualAverage))
        {
          equalAverageValue = "1";
        }

        SetXmlNodeString(
          ExcelConditionalFormattingConstants.Paths.EqualAverageAttribute,
          equalAverageValue,
          true);
      }
    }

    /// <summary>
    /// Bottom attribute
    /// </summary>
    internal protected bool? Bottom
    {
      get
      {
        bool? bottom = GetXmlNodeBoolNullable(
          ExcelConditionalFormattingConstants.Paths.BottomAttribute);

        // Bottom if TRUE
        return (bottom == true);
      }
      set
      {
        string bottomValue = string.Empty;

        // Only the types that needs the @Bottom
        if ((_type == eExcelConditionalFormattingRuleType.Bottom)
          || (_type == eExcelConditionalFormattingRuleType.BottomPercent))
        {
          bottomValue = "1";
        }

        SetXmlNodeString(
          ExcelConditionalFormattingConstants.Paths.BottomAttribute,
          bottomValue,
          true);
      }
    }

    /// <summary>
    /// Percent attribute
    /// </summary>
    internal protected bool? Percent
    {
      get
      {
        bool? percent = GetXmlNodeBoolNullable(
          ExcelConditionalFormattingConstants.Paths.PercentAttribute);

        // Bottom if TRUE
        return (percent == true);
      }
      set
      {
        string percentValue = string.Empty;

        // Only the types that needs the @Bottom
        if ((_type == eExcelConditionalFormattingRuleType.BottomPercent)
          || (_type == eExcelConditionalFormattingRuleType.TopPercent))
        {
          percentValue = "1";
        }

        SetXmlNodeString(
          ExcelConditionalFormattingConstants.Paths.PercentAttribute,
          percentValue,
          true);
      }
    }

    /// <summary>
    /// TimePeriod
    /// </summary>
    internal protected eExcelConditionalFormattingTimePeriodType TimePeriod
    {
      get
      {
        return ExcelConditionalFormattingTimePeriodType.GetTypeByAttribute(
          GetXmlNodeString(ExcelConditionalFormattingConstants.Paths.TimePeriodAttribute));
      }
      set
      {
        SetXmlNodeString(
          ExcelConditionalFormattingConstants.Paths.TimePeriodAttribute,
          ExcelConditionalFormattingTimePeriodType.GetAttributeByType(value),
          true);
      }
    }

    /// <summary>
    /// Operator
    /// </summary>
    internal protected eExcelConditionalFormattingOperatorType Operator
    {
      get
      {
        return ExcelConditionalFormattingOperatorType.GetTypeByAttribute(
          GetXmlNodeString(ExcelConditionalFormattingConstants.Paths.OperatorAttribute));
      }
      set
      {
        SetXmlNodeString(
          ExcelConditionalFormattingConstants.Paths.OperatorAttribute,
          ExcelConditionalFormattingOperatorType.GetAttributeByType(value),
          true);
      }
    }

    /// <summary>
    /// Formula
    /// </summary>
    public string Formula
    {
      get
      {
        return GetXmlNodeString(
          ExcelConditionalFormattingConstants.Paths.Formula);
      }
      set
      {
        SetXmlNodeString(
          ExcelConditionalFormattingConstants.Paths.Formula,
          value);
      }
    }

    /// <summary>
    /// Formula2
    /// </summary>
    public string Formula2
    {
      get
      {
        return GetXmlNodeString(
          string.Format(
            "{0}[position()=2]",
          // {0}
            ExcelConditionalFormattingConstants.Paths.Formula));
      }
      set
      {
        // Create/Get the first <formula> node (ensure that it exists)
        var firstNode = CreateComplexNode(
          TopNode,
          string.Format(
            "{0}[position()=1]",
          // {0}
            ExcelConditionalFormattingConstants.Paths.Formula));

        // Create/Get the seconde <formula> node (ensure that it exists)
        var secondNode = CreateComplexNode(
          TopNode,
          string.Format(
            "{0}[position()=2]",
          // {0}
            ExcelConditionalFormattingConstants.Paths.Formula));

        // Save the formula in the second <formula> node
        secondNode.InnerText = value;
      }
    }

    private ExcelConditionalFormattingAsType _as = null;
    /// <summary>
    /// Provides access to type conversion for all conditional formatting rules.
    /// </summary>
    public ExcelConditionalFormattingAsType As 
    { 
        get
        {
            if(_as==null)
            {
                _as = new ExcelConditionalFormattingAsType(this);
            }
            return _as;
        }
    } 
    #endregion Internal Properties
    /****************************************************************************************/
    internal protected void SetStyle(ExcelDxfStyleConditionalFormatting style)
    {
       _style = style;
       DxfId = int.MinValue;
    }
  }
}