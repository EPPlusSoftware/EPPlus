using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Table.PivotTable.Calculation;
using OfficeOpenXml.Table.PivotTable.Calculation.Functions;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
namespace EPPlusTest.Table.PivotTable.Calculation
{
    [TestClass]
	public class VerifyPivotCalculationWorkbookTests : TestBase
	{
		static ExcelPackage _package;
		static ExcelWorksheet _ptWs;
		static ExcelWorksheet _ptWs2;
		[ClassInitialize]
		public static void Init(TestContext context)


		{
			InitBase();
			_package = OpenTemplatePackage("PivotTableCalculation.xlsx");
			_ptWs = _package.Workbook.Worksheets["PivotTables"];
		}
		[ClassCleanup]
		public static void Cleanup()
		{
            SaveAndCleanup(_package);
            _package.Dispose();
		}
        [TestMethod]
        public void VerifyCalculationPivotTable1()
        {
            var pt = _ptWs.PivotTables["PivotTable1"];

            Assert.AreEqual(253396.22, GetPtData(pt, 0, "Australia", "TRUE"));
            Assert.AreEqual(24.581, GetPtData(pt, 1, "Australia", "FALSE"));
            
            Assert.AreEqual(0.025, (double)GetPtData(pt, 2, "Australia", null), 0.00000001D);

            Assert.AreEqual(356879.28, GetPtData(pt, 0, "Peru", "true"));
            Assert.AreEqual(14.3445, (double)GetPtData(pt, 1, "Peru", "true"), 0.0000001D);
            Assert.AreEqual(0.03, (double)GetPtData(pt, 2, "Peru", "true"), 0.0000001D);

            Assert.AreEqual(8996331.09, GetPtData(pt, 0, null, null));
            Assert.AreEqual(18.639405, (double)GetPtData(pt, 1, null, null), 0.00000001D);
            Assert.AreEqual(1D, (double)GetPtData(pt, 2, null, null), 0.00000001D);
        }
        [TestMethod]
        public void VerifyCalculationPivotTable2()
        {
            var pt = _ptWs.PivotTables["PivotTable2"];

            Assert.AreEqual(49286.72, GetPtData(pt, 0, "Austria", "Niedersachsen", "TRUE"));
            Assert.AreEqual(54215D, GetPtData(pt, 1, "Austria", "Niedersachsen", "TRUE"));

            Assert.AreEqual(117336.43, GetPtData(pt, 0, "Belgium", null, "False"));
            Assert.AreEqual(129070D, GetPtData(pt, 1, "Belgium", null, "false"));

            Assert.AreEqual(8996331.09, GetPtData(pt, 0, null, null, null));
            Assert.AreEqual(9895964D, GetPtData(pt, 1, null, null, null));
        }
        [TestMethod]
        public void VerifyCalculationPivotTable3()
        {
            var pt = _ptWs.PivotTables["PivotTable3"];

            Assert.AreEqual(8996331.09, GetPtData(pt, 0));
            Assert.AreEqual(3727.881, GetPtData(pt, 1));
            Assert.AreEqual(9689.13, GetPtData(pt, 2));
            Assert.AreEqual(9895964.00, GetPtData(pt, 3));

        }
        [TestMethod]
        public void VerifyCalculationPivotTable4()
        {
            var pt = _ptWs.PivotTables["PivotTable4"];
            //Santa Catarina	134091,14	44697,04667	3

            Assert.AreEqual(134091.14, GetPtData(pt, 0, "Santa Catarina"));
            Assert.AreEqual(44697.04667, (double)GetPtData(pt, 1, "Santa Catarina"), 0.00001);
            Assert.AreEqual(3D, GetPtData(pt, 2, "Santa Catarina"));
            Assert.AreEqual(8996331.09, GetPtData(pt, 0, null));
            Assert.AreEqual(44981.65545, GetPtData(pt, 1, null));
            Assert.AreEqual(200D, GetPtData(pt, 2, null));
        }
        [TestMethod]
        public void VerifyCalculationPivotTable5()
        {
            var pt = _ptWs.PivotTables["PivotTable5"];

            //Collapsed country with SubTotal Function - None.
            Assert.AreEqual(ErrorValues.RefError, GetPtData(pt, 0, "Australia", "Pskov Oblast"));
            Assert.AreEqual(273798.42, GetPtData(pt, 0, "Australia", null));

            //Expanded country with SubTotal Function - None.
            Assert.AreEqual(50879.73, GetPtData(pt, 0, "Belgium", "Rogaland"));
            Assert.AreEqual(ErrorValues.RefError, GetPtData(pt, 0, "Belgium", null));
        }
        [TestMethod]
        public void VerifyCalculationPivotTable6()
        {
            var pt = _ptWs.PivotTables["PivotTable6"];

            //Canada Sum	280213,09 Canada Count	8 Canada Average	35026,63625
            
            Assert.AreEqual(280213.09, GetPtData(pt, 0, null, "Country[Canada;Sum]"));
            Assert.AreEqual(8D, GetPtData(pt, 0, null, "Country[Canada,Count]"));
            Assert.AreEqual(35026.63625, GetPtData(pt, 0, null, "Country[Canada,Average]"));        
        }

        [TestMethod]
        public void VerifyCalculationPivotTable7()
        {
            var ws = _package.Workbook.Worksheets["PivotTableMultSubtotals"];
            var pt = ws.PivotTables["PivotTable7"];
            ws.Calculate();

            Assert.AreEqual(33997.99, (double)ws.Cells["A2"].Value);
            Assert.AreEqual(33997.99, (double)ws.Cells["A3"].Value);
            Assert.AreEqual(1D, (double)ws.Cells["A4"].Value);
            Assert.AreEqual(0D, (double)ws.Cells["A5"].Value);
            Assert.AreEqual(ErrorValues.RefError, ws.Cells["A6"].Value);
        }

        private object GetPtData(ExcelPivotTable pt, int datafield, params object[] values)
		{
			var l = new List<PivotDataFieldItemSelection>();
			int ix = 0;
			foreach (var f in pt.RowColumnFieldIndicies)
			{				
                if (values!=null && values[ix] != null)
				{					
                    if(values[ix].ToString().Contains("["))
                    {
                        var tokens = SourceCodeTokenizer.Default.Tokenize(values[ix].ToString());
                        if(tokens.Count==4)
                        {
                            var fieldTokens = SourceCodeTokenizer.Default.Tokenize(tokens[2].Value);
                            if(GetSubTotalFunctionFromString(fieldTokens[2].Value, out eSubTotalFunctions functions))
                            {
                                l.Add(new PivotDataFieldItemSelection(tokens[0].Value, fieldTokens[0].Value, functions));
                            }
                            else
                            {
                                return ErrorValues.RefError;
                            }                            
                        }
                        else
                        {
                            return ErrorValues.RefError;
                        }
                    }
                    else
                    {
                        l.Add(new PivotDataFieldItemSelection(pt.Fields[f].Name, values[ix]));
                    }
                }
				ix++;
			}

			return pt.GetPivotData(pt.DataFields[datafield].Name, l);
		}

        private bool GetSubTotalFunctionFromString(string value, out eSubTotalFunctions function)
        {
            switch(value.ToLower())
            {
                case "sum":
                    function = eSubTotalFunctions.Sum;
                    break;
                case "count":
                    function = eSubTotalFunctions.CountA;
                    break;
                case "count nums":
                    function = eSubTotalFunctions.Count;
                    break;
                case "average":
                    function = eSubTotalFunctions.Avg;
                    break;
                case "min":
                    function = eSubTotalFunctions.Min;
                    break;
                case "max":
                    function = eSubTotalFunctions.Max;
                    break;
                case "stddev":
                    function = eSubTotalFunctions.StdDev;
                    break;
                case "stddevp":
                    function = eSubTotalFunctions.StdDevP;
                    break;
                case "var":
                    function = eSubTotalFunctions.Var;
                    break;
                case "varp":
                    function = eSubTotalFunctions.VarP;
                    break;
                case "product":
                    function = eSubTotalFunctions.Product;
                    break;
                default:
                    function = eSubTotalFunctions.None;
                    return false;
            }
            return true;
        }
    }
}
