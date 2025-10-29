using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Data.Connection;
using System;
using System.Data.Common;

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
        public void AddDatabaseConnectionSimpeTest()
        {
            var dbConn = _pck.Workbook.Connections.AddDatabase("Connection1", "DRIVER=SQL Server;SERVER=epplusprod.database.windows.net;UID=epplusreadonly;APP=Microsoft Office;WSID=JANNESTHINKPAD;DATABASE=master");
            dbConn.DatabaseProperties.Command = "SELECT * FROM master.sys.all_columns all_columns";
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
            using(var p=OpenTemplatePackage("PowerQueryConnection.xlsx"))
            {
                foreach(var c in p.Workbook.Connections)
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
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
            _pck.Dispose();
        }
    }
}
