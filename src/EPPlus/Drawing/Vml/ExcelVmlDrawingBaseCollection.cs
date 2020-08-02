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

namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// Base collection for VML drawings
    /// </summary>
    public class ExcelVmlDrawingBaseCollection
    {        
        internal ExcelVmlDrawingBaseCollection(ExcelPackage pck, ExcelWorksheet ws, Uri uri)
        {
            VmlDrawingXml = new XmlDocument();
            VmlDrawingXml.PreserveWhitespace = false;
            
            NameTable nt=new NameTable();
            NameSpaceManager = new XmlNamespaceManager(nt);
            NameSpaceManager.AddNamespace("v", ExcelPackage.schemaMicrosoftVml);
            NameSpaceManager.AddNamespace("o", ExcelPackage.schemaMicrosoftOffice);
            NameSpaceManager.AddNamespace("x", ExcelPackage.schemaMicrosoftExcel);
            Uri = uri;
            if (uri == null)
            {
                Part = null;
            }
            else
            {
                Part=pck.ZipPackage.GetPart(uri);
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
