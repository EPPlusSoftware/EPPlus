/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
*************************************************************************************************
 Date               Author                   Change
*************************************************************************************************
 02/03/2020         EPPlus Software AB       Added
*************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.Core.Worksheet
{
    internal static class WorksheetRangeCommonHelper
    {
        internal static void AdjustDvAndCfFormulasRow(ExcelWorksheet ws, int rowFrom, int rows)
        {
            for (int i = 0; i < ws.DataValidations.Count; i++)
            {
                var type = ws.DataValidations.GetFormulas(ws.DataValidations[i], out IExcelDataValidationFormula Formula, out IExcelDataValidationFormula Formula2);

                if (Formula != null)
                {
                    if(Formula.ExcelFormula != null)
                    {
                        Formula.ExcelFormula = ExcelCellBase.UpdateFormulaReferences(Formula.ExcelFormula, rows, 0, rowFrom, 0, ws.Name, ws.Name);
                    }

                    if (Formula2 != null)
                    {
                        if (Formula2.ExcelFormula != null)
                        {
                            Formula2.ExcelFormula = ExcelCellBase.UpdateFormulaReferences(Formula2.ExcelFormula, rows, 0, rowFrom, 0, ws.Name, ws.Name);
                        }
                    }
                }
            }

            foreach (ExcelConditionalFormattingRule cf in ws.ConditionalFormatting)
            {
                UpdateCFformulaReferences(cf, rows, 0, rowFrom, 0, ws.Name);
            }
        }

        internal static void AdjustDvAndCfFormulasColumn(ExcelWorksheet ws, int columnFrom, int columns)
        {
            foreach (ExcelDataValidation dv in ws.DataValidations)
            {
                if (dv is ExcelDataValidationWithFormula<IExcelDataValidationFormula> dvFormula)
                {
                    dvFormula.Formula.ExcelFormula = ExcelCellBase.UpdateFormulaReferences(dvFormula.Formula.ExcelFormula, 0, columns, 0, columnFrom, ws.Name, ws.Name);
                    if (dv is ExcelDataValidationWithFormula2<IExcelDataValidationFormula> dvFormula2)
                    {
                        dvFormula2.Formula2.ExcelFormula = ExcelCellBase.UpdateFormulaReferences(dvFormula2.Formula2.ExcelFormula, 0, columns, 0, columnFrom, ws.Name, ws.Name);
                    }
                }
            }

            foreach (ExcelConditionalFormattingRule cf in ws.ConditionalFormatting)
            {
                UpdateCFformulaReferences(cf, 0, columns, 0, columnFrom, ws.Name);
            }
        }

        internal static void UpdateCFformulaReferences(ExcelConditionalFormattingRule cf, int rows, int columns, int rowFrom, int columnFrom, string currentSheet)
        {
            if (cf is ExcelConditionalFormattingTwoColorScale)
            {
                var colorScale = cf.As.TwoColorScale;
                if (colorScale.LowValue.Formula != null)
                {
                    colorScale.LowValue.Formula = ExcelCellBase.UpdateFormulaReferences(colorScale.LowValue.Formula, rows, columns, rowFrom, columnFrom, currentSheet, currentSheet);
                }
                if (colorScale.HighValue.Formula != null)
                {
                    colorScale.HighValue.Formula = ExcelCellBase.UpdateFormulaReferences(colorScale.HighValue.Formula, rows, columns, rowFrom, columnFrom, currentSheet, currentSheet);
                }
                if (cf is ExcelConditionalFormattingThreeColorScale)
                {
                    var threeColorScale = cf.As.ThreeColorScale;
                    if (threeColorScale.MiddleValue.Formula != null)
                    {
                        threeColorScale.MiddleValue.Formula = ExcelCellBase.UpdateFormulaReferences(threeColorScale.MiddleValue.Formula, rows, columns, rowFrom, columnFrom, currentSheet, currentSheet);
                    }
                }
            }

            if (cf.IsIconSet)
            {
                ExcelConditionalFormattingIconDataBarValue[] iconArray = null;
                switch (cf.Type)
                {
                    case eExcelConditionalFormattingRuleType.ThreeIconSet:
                        var iconSet3 = (ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting3IconsSetType>)cf;
                        iconArray = iconSet3.GetIconArray();
                        break;
                    case eExcelConditionalFormattingRuleType.FourIconSet:
                        var iconSet4 = (ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting3IconsSetType>)cf;
                        iconArray = iconSet4.GetIconArray();
                        break;
                    case eExcelConditionalFormattingRuleType.FiveIconSet:
                        var iconSet5 = (ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting3IconsSetType>)cf;
                        iconArray = iconSet5.GetIconArray();
                        break;
                }

                for (int i = 0; i < iconArray.Length; i++)
                {
                    if (iconArray[i].Formula != null)
                    {
                        iconArray[i].Formula = ExcelCellBase.UpdateFormulaReferences(cf.Formula, rows, columns, rowFrom, columnFrom, currentSheet, currentSheet);
                    }
                }
            }

            if (!string.IsNullOrEmpty(cf.Formula))
            {
                cf.Formula = ExcelCellBase.UpdateFormulaReferences(cf.Formula, rows, columns, rowFrom, columnFrom, currentSheet, currentSheet);
            }
            if (!string.IsNullOrEmpty(cf.Formula2))
            {
                cf.Formula2 = ExcelCellBase.UpdateFormulaReferences(cf.Formula2, rows, columns, rowFrom, columnFrom, currentSheet, currentSheet);
            }
        }
    }
}