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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.Interfaces
{
    internal class HashInfo
    {
        public HashInfo(string relId)
        {
            RelId = relId;
        }
        public string RelId { get; set; }
        public int RefCount { get; set; }
    }
    internal interface IPictureRelationDocument
    {
        ExcelPackage Package { get; }
        Dictionary<string, HashInfo> Hashes { get; }
        ZipPackagePart RelatedPart { get; }
        Uri RelatedUri { get; }
    }
    internal interface IPictureContainer
    {
        IPictureRelationDocument RelationDocument { get; }
        string ImageHash { get; set; }
        Uri UriPic { get; set; }
        Packaging.ZipPackageRelationship RelPic { get; set; }
    }
}
