﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Constants
{
    internal static class Relationsships
    {
        internal const string schemaMetadata = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
        // Richdata (used in worksheet.sortstate)
        internal const string schemaRichDataValueStructureRelationship = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure";
        internal const string schemaRichDataValueTypeRelationship = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes";
        internal const string schemaRichDataValueRelationship = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue";
        internal const string schemaRichDataRelRelationship = "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel";

    }
}
