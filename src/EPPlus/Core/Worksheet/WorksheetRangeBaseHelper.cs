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
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace OfficeOpenXml.Core.Worksheet
{
    internal static class WorksheetRangeCommonHelper
    {
        internal static void AdjustDvAndCfFormulasRow(ExcelWorksheet ws, int rowFrom, int rows)
        {
            foreach (var dv in ws.DataValidations)
            {
                if (dv is ExcelDataValidationWithFormula<IExcelDataValidationFormula> dvFormula)
                {
                    dvFormula.Formula.ExcelFormula = ExcelCellBase.UpdateFormulaReferences(dvFormula.Formula.ExcelFormula, rows, 0, rowFrom, 0, ws.Name, ws.Name);
                    if (dv is ExcelDataValidationWithFormula2<IExcelDataValidationFormula> dvFormula2)
                    {
                        dvFormula2.Formula2.ExcelFormula = ExcelCellBase.UpdateFormulaReferences(dvFormula2.Formula2.ExcelFormula, rows, 0, rowFrom, 0, ws.Name, ws.Name);
                    }
                }
            }

            foreach (ExcelConditionalFormattingRule cf in ws.ConditionalFormatting)
            {
                if (!string.IsNullOrEmpty(cf.Formula))
                {
                    cf.Formula = ExcelCellBase.UpdateFormulaReferences(cf.Formula, rows, 0, rowFrom, 0, ws.Name, ws.Name);
                }
                if (!string.IsNullOrEmpty(cf.Formula2))
                {
                    cf.Formula2 = ExcelCellBase.UpdateFormulaReferences(cf.Formula2, rows, 0, rowFrom, 0, ws.Name, ws.Name);
                }
            }
        }
        internal static void AdjustDvAndCfFormulasDelete(ExcelRangeBase range, ExcelAddressBase affectedRange, eShiftTypeDelete shift)
        {
            var ws=range.Worksheet;
            foreach (var dv in ws.DataValidations)
            {
                if (dv is ExcelDataValidationWithFormula<IExcelDataValidationFormula> dvFormula)
                {
                    dvFormula.Formula.ExcelFormula = ExcelCellBase.UpdateFormulaReferences(dvFormula.Formula.ExcelFormula, range, affectedRange, shift, ws.Name, ws.Name);
                    if (dv is ExcelDataValidationWithFormula2<IExcelDataValidationFormula> dvFormula2)
                    {
                        dvFormula2.Formula2.ExcelFormula = ExcelCellBase.UpdateFormulaReferences(dvFormula2.Formula2.ExcelFormula, range, affectedRange, shift, ws.Name, ws.Name);
                    }
                }
            }

            foreach (ExcelConditionalFormattingRule cf in ws.ConditionalFormatting)
            {
                if (!string.IsNullOrEmpty(cf.Formula))
                {
                    cf.Formula = ExcelCellBase.UpdateFormulaReferences(cf.Formula, range, affectedRange, shift, ws.Name, ws.Name);
                }
                if (!string.IsNullOrEmpty(cf.Formula2))
                {
                    cf.Formula2 = ExcelCellBase.UpdateFormulaReferences(cf.Formula2, range, affectedRange, shift, ws.Name, ws.Name);
                }
            }
        }
        internal static void AdjustDvAndCfFormulasInsert(ExcelRangeBase range, ExcelAddressBase affectedRange, eShiftTypeInsert shift)
        {
            var ws = range.Worksheet;
            foreach (var dv in ws.DataValidations)
            {
                if (dv is ExcelDataValidationWithFormula<IExcelDataValidationFormula> dvFormula)
                {
                    dvFormula.Formula.ExcelFormula = ExcelCellBase.UpdateFormulaReferences(dvFormula.Formula.ExcelFormula, range, affectedRange, shift, ws.Name, ws.Name);
                    if (dv is ExcelDataValidationWithFormula2<IExcelDataValidationFormula> dvFormula2)
                    {
                        dvFormula2.Formula2.ExcelFormula = ExcelCellBase.UpdateFormulaReferences(dvFormula2.Formula2.ExcelFormula, range, affectedRange, shift, ws.Name, ws.Name);
                    }
                }
            }

            foreach (ExcelConditionalFormattingRule cf in ws.ConditionalFormatting)
            {
                if (!string.IsNullOrEmpty(cf.Formula))
                {
                    cf.Formula = ExcelCellBase.UpdateFormulaReferences(cf.Formula, range, affectedRange, shift, ws.Name, ws.Name);
                }
                if (!string.IsNullOrEmpty(cf.Formula2))
                {
                    cf.Formula2 = ExcelCellBase.UpdateFormulaReferences(cf.Formula2, range, affectedRange, shift, ws.Name, ws.Name);
                }
            }
        }

        internal static void AdjustDvAndCfFormulasColumn(ExcelWorksheet ws, int columnFrom, int columns)
        {
            foreach (var dv in ws.DataValidations)
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
                if (!string.IsNullOrEmpty(cf.Formula))
                {
                    cf.Formula = ExcelCellBase.UpdateFormulaReferences(cf.Formula, 0, columns, 0, columnFrom, ws.Name, ws.Name);
                }
                if (!string.IsNullOrEmpty(cf.Formula2))
                {
                    cf.Formula2 = ExcelCellBase.UpdateFormulaReferences(cf.Formula2, 0, columns, 0, columnFrom, ws.Name, ws.Name);
                }
            }
        }

    }
}