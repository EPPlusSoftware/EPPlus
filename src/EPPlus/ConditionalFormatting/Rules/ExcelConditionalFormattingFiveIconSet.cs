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
  07/07/2023         EPPlus Software AB       Epplus 7
 *************************************************************************************************/
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.ConditionalFormatting.Rules
{
    internal class ExcelConditionalFormattingFiveIconSet : 
        ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting5IconsSetType>, 
        IExcelConditionalFormattingFiveIconSet
    {
        internal ExcelConditionalFormattingFiveIconSet(
        ExcelAddress address,
        int priority,
        ExcelWorksheet worksheet)
            : base(
              eExcelConditionalFormattingRuleType.FiveIconSet,
              address,
              priority,
              worksheet)
        {
            Icon4 = CreateIcon(60, eExcelConditionalFormattingRuleType.FiveIconSet);
            Icon5 = CreateIcon(80, eExcelConditionalFormattingRuleType.FiveIconSet);
        }

        internal ExcelConditionalFormattingFiveIconSet(
            ExcelAddress address,
            int priority,
            ExcelWorksheet worksheet,
            bool stopIfTrue,
            XmlReader xr)
            : base(
            eExcelConditionalFormattingRuleType.FiveIconSet,
            address,
            priority,
            worksheet,
            stopIfTrue,
            xr)
        {
            Icon4 = CreateIcon(60, eExcelConditionalFormattingRuleType.FiveIconSet);
            Icon5 = CreateIcon(80, eExcelConditionalFormattingRuleType.FiveIconSet);

            ReadIcon(Icon4, xr);

            xr.Read();

            ReadIcon(Icon5, xr);

            xr.Read();
            xr.Read();
        }

        internal ExcelConditionalFormattingFiveIconSet(ExcelConditionalFormattingFiveIconSet copy, ExcelWorksheet newWs = null) : base(copy, newWs)
        {
            Icon4 = copy.Icon4;
            Icon5 = copy.Icon5;
        }


        internal override ExcelConditionalFormattingIconDataBarValue[] GetIconArray(bool reversed = false)
        {
            return reversed ? [Icon5, Icon4, Icon3, Icon2, Icon1] : [Icon1, Icon2, Icon3, Icon4, Icon5];
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingFiveIconSet(this, newWs);
        }

        /// <summary>
        /// Icon 4 value
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue Icon4
        {
            get;
            internal set;
        }

        /// <summary>
        /// Icon 4 value
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue Icon5
        {
            get;
            internal set;
        }

        public override bool Custom
        {
            get
            {
                var ret = base.Custom;

                if (Icon4.CustomIcon != null || Icon5.CustomIcon != null)
                {
                    ret = true;
                }

                return ret;
            }
        }

        internal override bool IsExtLst
        {
            get
            {
                if (Custom)
                {
                    return true;
                }

                if (ExcelAddressBase.RefersToOtherWorksheet(Icon5.Formula, _ws.Name))
                {
                    return true;
                }

                return base.IsExtLst;
            }
        }
    }
}
