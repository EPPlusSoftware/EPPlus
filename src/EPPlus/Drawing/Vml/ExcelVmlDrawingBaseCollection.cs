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
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Collections;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// Base collection for VML drawings
    /// </summary>
    public class ExcelVmlDrawingBaseCollection
    {        
        protected internal ExcelPackage _package;
        protected internal ExcelWorksheet _ws;
        internal ExcelVmlDrawingBaseCollection(ExcelWorksheet ws, Uri uri, string worksheetRelIdPath)
        {
            VmlDrawingXml = new XmlDocument();
            VmlDrawingXml.PreserveWhitespace = false;
            
            NameTable nt=new NameTable();
            NameSpaceManager = new XmlNamespaceManager(nt);
            NameSpaceManager.AddNamespace("v", ExcelPackage.schemaMicrosoftVml);
            NameSpaceManager.AddNamespace("o", ExcelPackage.schemaMicrosoftOffice);
            NameSpaceManager.AddNamespace("x", ExcelPackage.schemaMicrosoftExcel);
            Uri = uri;
            _package = ws.Workbook._package;
            _ws = ws;
            if (uri == null)
            {
                var id = _ws.SheetId;
                Uri = XmlHelper.GetNewUri(_package.ZipPackage, @"/xl/drawings/vmlDrawing{0}.vml", ref id);

                Part = _package.ZipPackage.CreatePart(Uri, "application/vnd.openxmlformats-officedocument.vmlDrawing", _package.Compression);
                var rel = _ws.Part.CreateRelationship(UriHelper.GetRelativeUri(_ws.WorksheetUri, Uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");
                _ws.SetXmlNodeString(worksheetRelIdPath, rel.Id);
                RelId = rel.Id;
            }
            else
            {
                Part=_package.ZipPackage.GetPart(uri);
                XmlHelper.LoadXmlSafe(VmlDrawingXml, Part.GetStream()); 
            }
        }
        internal ExcelWorksheet Worksheet { get; set; }
        internal XmlDocument VmlDrawingXml { get; set; }
        internal Uri Uri { get; set; }
        internal string RelId { get; set; }
        internal Packaging.ZipPackagePart Part { get; set; }
        internal XmlNamespaceManager NameSpaceManager { get; set; }
    }
}
