using OfficeOpenXml.RichData.Structures.Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.Mappings
{
    internal static class RichValueRelationMappings
    {
        public static string GetSchema(string richValueRelationKey)
        {
            if(richValueRelationKey == StructureKeyNames.LocalImages.Image.RelLocalImageIdentifier)
            {
                return ExcelPackage.schemaImage;
            }
            else
            {
                throw new ArgumentException($"{richValueRelationKey} is an unsupported rich value relation");
            }
        }
    }
}
