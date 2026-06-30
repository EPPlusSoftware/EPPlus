using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;

using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

/// <summary>
/// USDOLLAR is a legacy function preserved for Lotus 1-2-3 compatibility.
/// 
/// Microsoft's documentation claims this function "always shows U.S. currency",
/// but in practice Excel desktop renders it using the system's current locale,
/// identical to DOLLAR. For example, USDOLLAR(3) on a Swedish system produces
/// "3,00 kr", not "$3.00". Verified against Excel desktop output.
/// 
/// We match the actual Excel behavior, not the documentation.
/// </summary>
[FunctionMetadata(
   Category = ExcelFunctionCategory.Text,
   EPPlusVersion = "8.6",
   Description = "Legacy Lotus 1-2-3 compatibility function. Behaves identically to DOLLAR.",
   SupportsArrays = true)]
internal class UsDollar : Dollar
{
}