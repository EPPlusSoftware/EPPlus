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
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// Conditional formatting with a five icon set
    /// </summary>
    public class ExcelConditionalFormattingFiveIconSet
    : ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting5IconsSetType>, IExcelConditionalFormattingFiveIconSet
  {
    /****************************************************************************************/

    #region Private Properties

    #endregion Private Properties

    /****************************************************************************************/

    #region Constructors
    /// <summary>
    /// 
    /// </summary>
    /// <param name="priority"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode"></param>
    /// <param name="namespaceManager"></param>
    internal ExcelConditionalFormattingFiveIconSet(
      ExcelAddress address,
      int priority,
      ExcelWorksheet worksheet,
      XmlNode itemElementNode,
      XmlNamespaceManager namespaceManager)
      : base(
        eExcelConditionalFormattingRuleType.FiveIconSet,
        address,
        priority,
        worksheet,
        itemElementNode,
        (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
    {
        if (itemElementNode != null && itemElementNode.HasChildNodes)
        {
            XmlNode iconNode4 = TopNode.SelectSingleNode("d:iconSet/d:cfvo[position()=4]", NameSpaceManager);
            Icon4 = new ExcelConditionalFormattingIconDataBarValue(
                    eExcelConditionalFormattingRuleType.FiveIconSet,
                    address,
                    worksheet,
                    iconNode4,
                    namespaceManager);
            
            XmlNode iconNode5 = TopNode.SelectSingleNode("d:iconSet/d:cfvo[position()=5]", NameSpaceManager);
            Icon5 = new ExcelConditionalFormattingIconDataBarValue(
                    eExcelConditionalFormattingRuleType.FiveIconSet,
                    address,
                    worksheet,
                    iconNode5,
                    namespaceManager);
        }
        else
        {
            XmlNode iconSetNode = TopNode.SelectSingleNode("d:iconSet", NameSpaceManager);
            var iconNode4 = iconSetNode.OwnerDocument.CreateElement(ExcelConditionalFormattingConstants.Paths.Cfvo, ExcelPackage.schemaMain);
            iconSetNode.AppendChild(iconNode4);

            Icon4 = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Percent,
                    60,
                    "",
                    eExcelConditionalFormattingRuleType.ThreeIconSet,
                    address,
                    priority,
                    worksheet,
                    iconNode4,
                    namespaceManager);

            var iconNode5 = iconSetNode.OwnerDocument.CreateElement(ExcelConditionalFormattingConstants.Paths.Cfvo, ExcelPackage.schemaMain);
            iconSetNode.AppendChild(iconNode5);

            Icon5 = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Percent,
                    80,
                    "",
                    eExcelConditionalFormattingRuleType.ThreeIconSet,
                    address,
                    priority,
                    worksheet,
                    iconNode5,
                    namespaceManager);
        }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="priority"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode"></param>
    internal ExcelConditionalFormattingFiveIconSet(
      ExcelAddress address,
      int priority,
      ExcelWorksheet worksheet,
      XmlNode itemElementNode)
      : this(
        address,
        priority,
        worksheet,
        itemElementNode,
        null)
    {
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="priority"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    internal ExcelConditionalFormattingFiveIconSet(
      ExcelAddress address,
      int priority,
      ExcelWorksheet worksheet)
      : this(
        address,
        priority,
        worksheet,
        null,
        null)
    {
    }
        #endregion Constructors

        /// <summary>
        /// Icon 5 value
        /// </summary>
    public ExcelConditionalFormattingIconDataBarValue Icon5
    {
        get;
        internal set;
    }

        /// <summary>
        /// Icon 4 value
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue Icon4
    {
        get;
        internal set;
    }
  }
  }
