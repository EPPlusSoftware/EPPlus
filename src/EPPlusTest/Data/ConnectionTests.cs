using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Data.Connection;
using System;
using System.Data.Common;
using System.Text;

namespace EPPlusTest.Data
{
    [TestClass]

    public class ConnectionTests : TestBase
    {
        private static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("Connections.xlsx", true);
            _pck.Workbook.Worksheets.Add("Sheet1");
        }
        [TestMethod]
        public void AddTextSimpeTest()
        {
            var sourceFile = "C:\\kod\\EPPlusSoftware\\EPPlus\\src\\EPPlusTest\\Resources\\Textfiles\\FixedWidth_FileList.txt";
            var textConn = _pck.Workbook.Connections.AddText("Connection2", sourceFile);
            //textConn.TextProperties.Fields.Add(new OfficeOpenXml.Data.Connection.ExcelConnectionTextField(OfficeOpenXml.Data.Connection.eConnectionTextFieldType.Text));
        }
        [TestMethod]
        public void AddWebSimpeTest()
        {
            var webConn = _pck.Workbook.Connections.AddWeb("Connection3", "https://epplussoftware.com/en/LicenseOverview/");
            webConn.WebProperties.Tables.Add(new ExcelHtmlTableReference(1));         
        }
        [TestMethod]
        public void ReadWebPowerQuery()
        {
            using( var p=OpenTemplatePackage("PowerQueryConnection.xlsx"))
            {
                foreach(var c in p.Workbook.Connections)
                {

                }
                Assert.IsNotNull(p.Workbook.Connections.PowerQuerySettings);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ReadWebPowerQueryEPP()
        {
            using (var p = OpenTemplatePackage("PowerQueryConnectionEPP.xlsx"))
            {
                foreach (var c in p.Workbook.Connections)
                {

                }
                Assert.IsNotNull(p.Workbook.Connections.PowerQuerySettings);
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void AddWebPowerQuery()
        {
            var dbConn = _pck.Workbook.Connections.AddDatabase("PowerQueryDb", "Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=\"Table 6\";Extended Properties=\"\"");
            dbConn.DatabaseProperties.Command = "SELECT * FROM [Table 6]";
            dbConn.KeepAlive = true;

            var pcs = _pck.Workbook.Connections.PowerQuerySettings;
            pcs.Create();
            pcs.PowerQueryFormulas = "section Section1;\r\n\r\nshared #\"Table 6\" = let\r\n    Source = Web.BrowserContents(\"https://epplussoftware.com/en/LicenseOverview/\"),\r\n    #\"Extracted Table From Html\" = Html.Table(Source, {{\"Column1\", \"DIV[id='category-packageperpetual'] > DIV.row.period-length-item-24.content-container:nth-child(3) > DIV.col-11.border.border-light.mx-3.my-2.p-3.rounded.rounded-lg.bg-light.shadow > DIV.row > DIV.col-12 > DIV.row.d-flex.justify-content-between > DIV.col-sm-11.col-lg-5.bg-white.rounded.rounded-lg.border.border-secondary.ml-2.mr-3.pt-2.table-responsive > TABLE.table.table-sm.float-right > * > TR > TH:not([colspan]):not([rowspan]):nth-child(1):nth-last-child(2), DIV[id='category-packageperpetual'] > DIV.row.period-length-item-24.content-container:nth-child(3) > DIV.col-11.border.border-light.mx-3.my-2.p-3.rounded.rounded-lg.bg-light.shadow > DIV.row > DIV.col-12 > DIV.row.d-flex.justify-content-between > DIV.col-sm-11.col-lg-5.bg-white.rounded.rounded-lg.border.border-secondary.ml-2.mr-3.pt-2.table-responsive > TABLE.table.table-sm.float-right > * > TR > TD:not([colspan]):not([rowspan]):nth-child(1):nth-last-child(2), DIV[id='category-packageperpetual'] > DIV.row.period-length-item-24.content-container:nth-child(3) > DIV.col-11.border.border-light.mx-3.my-2.p-3.rounded.rounded-lg.bg-light.shadow > DIV.row > DIV.col-12 > DIV.row.d-flex.justify-content-between > DIV.col-sm-11.col-lg-5.bg-white.rounded.rounded-lg.border.border-secondary.ml-2.mr-3.pt-2.table-responsive > TABLE.table.table-sm.float-right > * > TR > TD[colspan=\"\"3\"\"]:not([rowspan]):nth-child(1):nth-last-child(1)\"}, {\"Column2\", \"DIV[id='category-packageperpetual'] > DIV.row.period-length-item-24.content-container:nth-child(3) > DIV.col-11.border.border-light.mx-3.my-2.p-3.rounded.rounded-lg.bg-light.shadow > DIV.row > DIV.col-12 > DIV.row.d-flex.justify-content-between > DIV.col-sm-11.col-lg-5.bg-white.rounded.rounded-lg.border.border-secondary.ml-2.mr-3.pt-2.table-responsive > TABLE.table.table-sm.float-right > * > TR > TH:not([colspan]):not([rowspan]):nth-child(1):nth-last-child(2) + TH:not([colspan]):not([rowspan]):nth-child(2):nth-last-child(1), DIV[id='category-packageperpetual'] > DIV.row.period-length-item-24.content-container:nth-child(3) > DIV.col-11.border.border-light.mx-3.my-2.p-3.rounded.rounded-lg.bg-light.shadow > DIV.row > DIV.col-12 > DIV.row.d-flex.justify-content-between > DIV.col-sm-11.col-lg-5.bg-white.rounded.rounded-lg.border.border-secondary.ml-2.mr-3.pt-2.table-responsive > TABLE.table.table-sm.float-right > * > TR > TD:not([colspan]):not([rowspan]):nth-child(1):nth-last-child(2) + TD:not([colspan]):not([rowspan]):nth-child(2):nth-last-child(1), DIV[id='category-packageperpetual'] > DIV.row.period-length-item-24.content-container:nth-child(3) > DIV.col-11.border.border-light.mx-3.my-2.p-3.rounded.rounded-lg.bg-light.shadow > DIV.row > DIV.col-12 > DIV.row.d-flex.justify-content-between > DIV.col-sm-11.col-lg-5.bg-white.rounded.rounded-lg.border.border-secondary.ml-2.mr-3.pt-2.table-responsive > TABLE.table.table-sm.float-right > * > TR > TD[colspan=\"\"3\"\"]:not([rowspan]):nth-child(1):nth-last-child(1)\"}, {\"Column3\", \"DIV[id='category-packageperpetual'] > DIV.row.period-length-item-24.content-container:nth-child(3) > DIV.col-11.border.border-light.mx-3.my-2.p-3.rounded.rounded-lg.bg-light.shadow > DIV.row > DIV.col-12 > DIV.row.d-flex.justify-content-between > DIV.col-sm-11.col-lg-5.bg-white.rounded.rounded-lg.border.border-secondary.ml-2.mr-3.pt-2.table-responsive > TABLE.table.table-sm.float-right > * > TR > TD[colspan=\"\"3\"\"]:not([rowspan]):nth-child(1):nth-last-child(1)\"}}, [RowSelector=\"DIV[id='category-packageperpetual'] > DIV.row.period-length-item-24.content-container:nth-child(3) > DIV.col-11.border.border-light.mx-3.my-2.p-3.rounded.rounded-lg.bg-light.shadow > DIV.row > DIV.col-12 > DIV.row.d-flex.justify-content-between > DIV.col-sm-11.col-lg-5.bg-white.rounded.rounded-lg.border.border-secondary.ml-2.mr-3.pt-2.table-responsive > TABLE.table.table-sm.float-right > * > TR\"]),\r\n    #\"Changed Type\" = Table.TransformColumnTypes(#\"Extracted Table From Html\",{{\"Column1\", type text}, {\"Column2\", type text}, {\"Column3\", type text}})\r\nin\r\n    #\"Changed Type\";";
            var mdXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><LocalPackageMetadataFile xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><Items><Item><ItemLocation><ItemType>AllFormulas</ItemType><ItemPath /></ItemLocation><StableEntries><Entry Type=\"Relationships\" Value=\"sAAAAAA==\" /></StableEntries></Item><Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/Table%206</ItemPath></ItemLocation><StableEntries><Entry Type=\"QueryID\" Value=\"s82ea1023-d0be-4d8f-8a0c-bc45dd8f2a43\" /><Entry Type=\"FillEnabled\" Value=\"l1\" /><Entry Type=\"FillObjectType\" Value=\"sTable\" /><Entry Type=\"FillToDataModelEnabled\" Value=\"l0\" /><Entry Type=\"IsPrivate\" Value=\"l0\" /><Entry Type=\"BufferNextRefresh\" Value=\"l1\" /><Entry Type=\"ResultType\" Value=\"sTable\" /><Entry Type=\"NameUpdatedAfterFill\" Value=\"l0\" /><Entry Type=\"FillTarget\" Value=\"sTable_6\" /><Entry Type=\"FilledCompleteResultToWorksheet\" Value=\"l1\" /><Entry Type=\"AddedToDataModel\" Value=\"l0\" /><Entry Type=\"FillCount\" Value=\"l3\" /><Entry Type=\"FillErrorCode\" Value=\"sUnknown\" /><Entry Type=\"FillErrorCount\" Value=\"l0\" /><Entry Type=\"FillLastUpdated\" Value=\"d2025-10-24T11:49:08.4038620Z\" /><Entry Type=\"FillColumnTypes\" Value=\"sBgYG\" /><Entry Type=\"FillColumnNames\" Value=\"s[&quot;Column1&quot;,&quot;Column2&quot;,&quot;Column3&quot;]\" /><Entry Type=\"FillStatus\" Value=\"sComplete\" /><Entry Type=\"RelationshipInfoContainer\" Value=\"s{&quot;columnCount&quot;:3,&quot;keyColumnNames&quot;:[],&quot;queryRelationships&quot;:[],&quot;columnIdentities&quot;:[&quot;Section1/Table 6/AutoRemovedColumns1.{Column1,0}&quot;,&quot;Section1/Table 6/AutoRemovedColumns1.{Column2,1}&quot;,&quot;Section1/Table 6/AutoRemovedColumns1.{Column3,2}&quot;],&quot;ColumnCount&quot;:3,&quot;KeyColumnNames&quot;:[],&quot;ColumnIdentities&quot;:[&quot;Section1/Table 6/AutoRemovedColumns1.{Column1,0}&quot;,&quot;Section1/Table 6/AutoRemovedColumns1.{Column2,1}&quot;,&quot;Section1/Table 6/AutoRemovedColumns1.{Column3,2}&quot;],&quot;RelationshipInfo&quot;:[]}\" /></StableEntries></Item><Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/Table%206/Source</ItemPath></ItemLocation><StableEntries /></Item><Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/Table%206/Extracted%20Table%20From%20Html</ItemPath></ItemLocation><StableEntries /></Item><Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/Table%206/Changed%20Type</ItemPath></ItemLocation><StableEntries /></Item><Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/Table%206%20(2)</ItemPath></ItemLocation><StableEntries><Entry Type=\"QueryID\" Value=\"sf373d4dc-19ad-40d0-afe9-f26f0647968f\" /><Entry Type=\"FillEnabled\" Value=\"l1\" /><Entry Type=\"FillObjectType\" Value=\"sTable\" /><Entry Type=\"FillToDataModelEnabled\" Value=\"l0\" /><Entry Type=\"IsPrivate\" Value=\"l0\" /><Entry Type=\"ResultType\" Value=\"sTable\" /><Entry Type=\"NameUpdatedAfterFill\" Value=\"l0\" /><Entry Type=\"FillTarget\" Value=\"sTable_6__2\" /><Entry Type=\"FilledCompleteResultToWorksheet\" Value=\"l1\" /><Entry Type=\"FillStatus\" Value=\"sComplete\" /><Entry Type=\"FillColumnNames\" Value=\"s[&quot;Column1&quot;,&quot;Column2&quot;,&quot;Column3&quot;]\" /><Entry Type=\"FillColumnTypes\" Value=\"sBgYG\" /><Entry Type=\"FillLastUpdated\" Value=\"d2025-10-30T14:08:56.2449821Z\" /><Entry Type=\"FillErrorCount\" Value=\"l0\" /><Entry Type=\"FillErrorCode\" Value=\"sUnknown\" /><Entry Type=\"FillCount\" Value=\"l3\" /><Entry Type=\"AddedToDataModel\" Value=\"l0\" /><Entry Type=\"LoadedToAnalysisServices\" Value=\"l0\" /><Entry Type=\"RelationshipInfoContainer\" Value=\"s{&quot;columnCount&quot;:3,&quot;keyColumnNames&quot;:[],&quot;queryRelationships&quot;:[],&quot;columnIdentities&quot;:[&quot;Section1/Table 6 (2)/AutoRemovedColumns1.{Column1,0}&quot;,&quot;Section1/Table 6 (2)/AutoRemovedColumns1.{Column2,1}&quot;,&quot;Section1/Table 6 (2)/AutoRemovedColumns1.{Column3,2}&quot;],&quot;ColumnCount&quot;:3,&quot;KeyColumnNames&quot;:[],&quot;ColumnIdentities&quot;:[&quot;Section1/Table 6 (2)/AutoRemovedColumns1.{Column1,0}&quot;,&quot;Section1/Table 6 (2)/AutoRemovedColumns1.{Column2,1}&quot;,&quot;Section1/Table 6 (2)/AutoRemovedColumns1.{Column3,2}&quot;],&quot;RelationshipInfo&quot;:[]}\" /><Entry Type=\"BufferNextRefresh\" Value=\"l1\" /></StableEntries></Item><Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/Table%206%20(2)/Source</ItemPath></ItemLocation><StableEntries /></Item><Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/Table%206%20(2)/Extracted%20Table%20From%20Html</ItemPath></ItemLocation><StableEntries /></Item><Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/Table%206%20(2)/Changed%20Type</ItemPath></ItemLocation><StableEntries /></Item></Items></LocalPackageMetadataFile>";
            pcs.MetadataXml.LoadXml(mdXml);
            var ws = _pck.Workbook.Worksheets[0];
            var tbl = ws.Tables.AddQueryTable(ws.Cells["H50:K52"], "Table_6", dbConn, ["Column1", "Column2", "Column3", "Formula"]);
            tbl.QueryTable.Fields[3].DataBoundColumn = false;
            tbl.Columns[3].SetFormula("Count(Table_6[[#This Row],[Column1]])");
            tbl.QueryTable.RefreshOnLoad = true;
            //var pt = ws.PivotTables.Add(ws.Cells["A5"], dbConn, "PivotTable1");
        }
        [TestMethod]
        public void ReadEPPWebPowerQuery()
        {
            using (var p = OpenTemplatePackage("ConnectionsEPPSaved.xlsx"))
            {
                foreach (var c in p.Workbook.Connections)
                {

                }
                Assert.IsNotNull(p.Workbook.Connections.PowerQuerySettings);
                SaveAndCleanup(p);
            }
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
            _pck.Dispose();
        }
    }
}
