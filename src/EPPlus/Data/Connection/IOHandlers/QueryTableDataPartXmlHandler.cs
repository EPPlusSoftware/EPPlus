using OfficeOpenXml.Constants;
using OfficeOpenXml.Core;
using OfficeOpenXml.Data.QueryTable;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils.EnumUtils;
using OfficeOpenXml.Utils.FileUtils;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeOpenXml.Data.Connection.IOHandlers
{
    internal class QueryTableDataPartXmlHandler : IDocumentPart<ExcelQueryTable>
    {
        ZipPackagePart Part { get; set; }
        public ExcelRangeBase DestinationRange { get; private set; }

        XmlHelper _xml;
        ExcelTable _table=null;
        ExcelWorksheet _ws=null;
        public QueryTableDataPartXmlHandler(ExcelTable table) 
        {
            _table = table;
            _ws = table.WorkSheet;
            var qtRels = _table.Part.GetRelationshipsByType(ExcelPackage.schemaRelationships + "/queryTable");
            if (qtRels.Any())
            {
                var rel = qtRels.First();
                CreateXmlHelper(table.WorkSheet._package, rel);
            }
        }
        public QueryTableDataPartXmlHandler(ExcelTable table, ExcelConnection connection, string[] fieldNames)
        {
            var zp = _ws._package.ZipPackage;
            int id = 1;
            Part = zp.CreatePart(XmlHelper.GetNewUri(zp, "/xl/queryTables/queryTable{0}.xml", ref id), ContentTypes.contentTypeQueryTable, CompressionLevel.Default);
            _table.Part.CreateRelationship(Part.Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/queryTable");
            XmlDocument xml = new XmlDocument();
            var sb = new StringBuilder();
            sb.Append($"<?xml version=\"1.0\" encoding=\"utf-8\"?><queryTable xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"  xmlns:xr16=\"http://schemas.microsoft.com/office/spreadsheetml/2017/revision16\" mc:Ignorable=\"xr16\"><queryTableRefresh nextId=\"{fieldNames.Length + 1}\"><queryTableFields count=\"{fieldNames.Length}\">");
            int fid=1;
            foreach (var fieldName in fieldNames)
            {
                sb.Append($"<id=\"{fid}\" name=\"{fieldName}\" tableColumnId=\"{fid}\" />");
                fid++;
            }
            xml.LoadXml("</queryTableFields></queryTableRefresh></queryTable>");
        }
        public QueryTableDataPartXmlHandler(ExcelWorksheet ws, ZipPackageRelationship rel)
        {
            _ws = ws;
            CreateXmlHelper(ws._package, rel);
        }
        private void CreateXmlHelper(ExcelPackage pck, ZipPackageRelationship rel)
        {
            var uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            Part = pck.ZipPackage.GetPart(uri);
            var xmlDoc = new XmlDocument();
            XmlHelper.LoadXmlSafe(xmlDoc, Part.GetStream());
            _xml = XmlHelperFactory.Create(pck.Workbook.NameSpaceManager, xmlDoc.DocumentElement);
            _xml.SchemaNodeOrder = ["queryTableRefresh", "queryTableFields", "queryTableDeletedFields", "sortState"];
        }
        public void Load(ExcelQueryTable item)
        {
            item.Name = _xml.GetXmlNodeString("@name");
            item.ConnectionId = _xml.GetXmlNodeInt("@connectionId");
            item.Headers = _xml.GetXmlNodeBool("@headers", true);
            item.RowNumbers = _xml.GetXmlNodeBool("@headers", false);
            item.DisableRefresh = _xml.GetXmlNodeBool("@disableRefresh", false);
            item.BackgroundRefresh = _xml.GetXmlNodeBoolNullable("@backgroundRefresh");
            item.FirstBackgroundRefresh = _xml.GetXmlNodeBoolNullable("firstBackgroundRefresh");
            item.GrowShrinkType = _xml.GetXmlEnum("@growShrinkType", QueryTableGrowShrinkType.InsertDelete);
            item.RefreshOnLoad = _xml.GetXmlNodeBool("@refreshOnLoad", false);
            item.FillFormulas = _xml.GetXmlNodeBool("@fillFormulas", false);
            item.RemoveDataOnSave = _xml.GetXmlNodeBoolNullable("@removeDataOnSave");
            item.DisableEdit = _xml.GetXmlNodeBool("@disableEdit", false);
            item.PreserveFormatting = _xml.GetXmlNodeBool("@preserveFormatting", true);
            item.AdjustColumnWidth = _xml.GetXmlNodeBool("@adjustColumnWidth", true);
            item.Intermediate = _xml.GetXmlNodeBool("@intermediate", false);

            item.AutoFormatId = _xml.GetXmlNodeIntNull("@autoFormatId");
            item.ApplyNumberFormats = _xml.GetXmlNodeBoolNullable("@applyNumberFormats");
            item.ApplyBorderFormats = _xml.GetXmlNodeBoolNullable("@applyBorderFormats");
            item.ApplyFontFormats = _xml.GetXmlNodeBoolNullable("@applyFontFormats");
            item.ApplyPatternFormats = _xml.GetXmlNodeBoolNullable("@applyPatternFormats");
            item.ApplyAlignmentFormats = _xml.GetXmlNodeBoolNullable("@applyAlignmentFormats");
            item.ApplyWidthHeightFormats = _xml.GetXmlNodeBoolNullable("@applyWidthHeightFormats");

            foreach (XmlElement node in _xml.GetNodes("d:queryTableRefresh/d:queryTableFields/d:queryTableField"))
            {
                var xh = XmlHelperFactory.Create(_xml.NameSpaceManager, node);
                var field = new ExcelQueryTableField()
                {
                    Id = xh.GetXmlNodeInt("@id"),
                    Name = xh.GetXmlNodeString("@name"),
                    TableColumnId = xh.GetXmlNodeInt("@tableColumnId", 0),
                    ClippedColumn = xh.GetXmlNodeBool("@clipperColumn"),
                    DataBoundColumn = xh.GetXmlNodeBool("@dataBound", true),
                    FillFormulaOnRefresh = xh.GetXmlNodeBool("@fillFormulas"),
                    RowNumbers = xh.GetXmlNodeBool("@rowNumbers")
                };
                item.Fields.Add(field);
            }
            foreach (XmlElement node in _xml.GetNodes("d:queryTableRefresh/d:queryTableDeletedFields/d:deletedField"))
            {
                item.DeletedFields.Add(node.Attributes[0].Value);
            }
            item.Connection = _ws.Workbook.Connections.FirstOrDefault(x => x.Id == item.ConnectionId);
            if (_table == null)
            {
                var name = _ws.Names[item.Name];
                if (name != null)
                {
                    DestinationRange = name;
                }
            }
            else
            {
                DestinationRange = _table.Range;
            }
        }
        public void Remove()
        {
            if (Part != null)
            {
                _table.WorkSheet.Workbook._package.ZipPackage.DeletePart(Part.Uri);
            }
        }
        public void Save(ExcelQueryTable item)
        {
            _xml.SetXmlNodeString("@name", item.Name);
            _xml.GetXmlNodeInt("@connectionId", item.ConnectionId);
            _xml.SetXmlNodeBool("@headers", item.Headers, true);
            _xml.SetXmlNodeBool("@headers", item.RowNumbers, false);
            _xml.SetXmlNodeBool("@disableRefresh", item.DisableRefresh, false);
            _xml.SetXmlNodeBoolNull("@backgroundRefresh", item.BackgroundRefresh);
            _xml.SetXmlNodeBoolNull("@firstBackgroundRefresh", item.FirstBackgroundRefresh);
            _xml.SetXmlNodeString("@growShrinkType", item.GrowShrinkType.ToEnumString(QueryTableGrowShrinkType.InsertDelete), true, false, true);
            _xml.SetXmlNodeBool("@refreshOnLoad", item.RefreshOnLoad, false);
            _xml.SetXmlNodeBool("@fillFormulas", item.FillFormulas, false);
            _xml.SetXmlNodeBoolNull("@removeDataOnSave", item.RemoveDataOnSave);
            _xml.SetXmlNodeBool("@disableEdit", item.DisableEdit, false);
            _xml.SetXmlNodeBool("@preserveFormatting", item.PreserveFormatting, true);
            _xml.SetXmlNodeBool("@adjustColumnWidth", item.AdjustColumnWidth, true);
            _xml.SetXmlNodeBool("@intermediate", item.Intermediate, false);

            _xml.SetXmlNodeInt("@autoFormatId", item.AutoFormatId);
            _xml.SetXmlNodeBoolNull("@applyNumberFormats", item.ApplyNumberFormats);
            _xml.SetXmlNodeBoolNull("@applyBorderFormats", item.ApplyBorderFormats);
            _xml.SetXmlNodeBoolNull("@applyFontFormats", item.ApplyFontFormats);
            _xml.SetXmlNodeBoolNull("@applyPatternFormats", item.ApplyPatternFormats);
            _xml.SetXmlNodeBoolNull("@applyAlignmentFormats", item.ApplyAlignmentFormats);
            _xml.SetXmlNodeBoolNull("@applyWidthHeightFormats", item.ApplyWidthHeightFormats);

            if(item.Fields.Count > 0)
            {
                var fn = (XmlElement)_xml.CreateNode("d:queryTableRefresh/d:queryTableFields");
                fn.RemoveAll();
                foreach (var f in item.Fields)
                {
                    var node = _xml.CreateNode("d:queryTableRefresh/d:queryTableFields/d:queryTableField", false, true);
                    var xh = XmlHelperFactory.Create(_xml.NameSpaceManager, node);
                    xh.SetXmlNodeInt("@id", f.Id);
                    xh.SetXmlNodeString("@name", f.Name);
                    xh.SetXmlNodeInt("@tableColumnId", f.TableColumnId, 0);
                    xh.SetXmlNodeBool("@clipped", f.ClippedColumn, false);
                    xh.SetXmlNodeBool("@dataBound", f.DataBoundColumn, true);
                    xh.SetXmlNodeBool("@fillFormulas", f.FillFormulaOnRefresh, false);
                    xh.SetXmlNodeBool("@rowNumbers", f.RowNumbers, false);
                }
                fn.SetAttribute("count", item.Fields.Count.ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                _xml.DeleteNode("d:queryTableRefresh/d:queryTableFields");
            }

            if (item.DeletedFields.Count > 0)
            {
                var fn = (XmlElement)_xml.CreateNode("d:queryTableRefresh/d:queryTableDeletedFields");
                fn.RemoveAll();

                foreach (var f in item.DeletedFields)
                {
                    var node = (XmlElement)_xml.CreateNode("d:queryTableRefresh/d:queryTableDeletedFields/d:deletedField", false, true);
                    node.SetAttribute("name", f);
                }
                fn.SetAttribute("count", item.DeletedFields.Count.ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                _xml.DeleteNode("d:queryTableRefresh/d:queryTableDeletedFields");
            }

            var qtXmlStream = Part.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write);
            _xml.TopNode.OwnerDocument.Save(qtXmlStream);
            qtXmlStream.Flush();
        }
    }
}
