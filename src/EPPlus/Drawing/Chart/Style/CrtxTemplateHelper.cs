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
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.Style
{
    internal class CrtxTemplateHelper
    {
        internal class RelationShipWithContent
        {
        }
        internal static void LoadCrtx(Stream stream, out XmlDocument chartXml, out XmlDocument styleXml, out XmlDocument colorsXml, out ZipPackagePart themePart, string fileName)
        {           
            chartXml = null;
            styleXml = null;
            colorsXml = null;
            themePart = null;
            try
            {
                ZipPackage p = new ZipPackage(stream);

                var chartRel = p.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument").FirstOrDefault();
                if (chartRel != null)
                {
                    var chartPart = p.GetPart(chartRel.TargetUri);
                    chartXml = new XmlDocument();
                    chartXml.Load(chartPart.GetStream());
                    var rels = chartPart.GetRelationships();
                    foreach (var rel in rels)
                    {
                        var part = p.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
                        switch (rel.RelationshipType)
                        {
                            case ExcelPackage.schemaThemeOverrideRelationships:
                                themePart = part;
                                break;
                            case ExcelPackage.schemaChartStyleRelationships:
                                styleXml = new XmlDocument();
                                styleXml.Load(part.GetStream());
                                break;
                            case ExcelPackage.schemaChartColorStyleRelationships:
                                colorsXml = new XmlDocument();
                                colorsXml.Load(part.GetStream());
                                break;
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                throw new InvalidDataException($"{fileName} has an invalid format", ex);
            }
        }
    }
}
