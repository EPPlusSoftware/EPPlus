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
    internal class ExcelConditionalFormattingFourIconSet : 
        ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting4IconsSetType>, 
        IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>
    {
        internal ExcelConditionalFormattingFourIconSet(
        ExcelAddress address,
        int priority,
        ExcelWorksheet worksheet)
            : base(
              eExcelConditionalFormattingRuleType.FourIconSet,
              address,
              priority,
              worksheet)
        {
            Icon4 = CreateIcon(75, eExcelConditionalFormattingRuleType.FourIconSet);
        }

        internal ExcelConditionalFormattingFourIconSet(
        ExcelAddress address,
        int priority,
        ExcelWorksheet worksheet,
        bool stopIfTrue,
        XmlReader xr)
            : base(
            eExcelConditionalFormattingRuleType.FourIconSet,
            address,
            priority,
            worksheet,
            stopIfTrue,
            xr)
        {
            Icon4 = CreateIcon(75, eExcelConditionalFormattingRuleType.FourIconSet);

            ReadIcon(Icon4, xr);

            xr.Read();
            xr.Read();
        }

        internal ExcelConditionalFormattingFourIconSet(ExcelConditionalFormattingFourIconSet copy) : base(copy)
        {
            Icon4 = copy.Icon4;
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingFourIconSet(this);
        }

        /// <summary>
        /// Icon 4 value
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue Icon4
        {
            get;
            internal set;
        }

        internal override bool IsExtLst
        {
            get
            {
                if (Icon1.CustomIcon != null ||
                    Icon2.CustomIcon != null ||
                    Icon3.CustomIcon != null ||
                    Icon4.CustomIcon != null)
                {
                    return true;
                }

                if (ExcelAddressBase.RefersToOtherWorksheet(Icon4.Formula, _ws.Name))
                {
                    return true;
                }

                return base.IsExtLst;
            }
        }
    }
}
