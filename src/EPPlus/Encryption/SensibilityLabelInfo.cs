/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial licenseXml to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/29/2024         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System.Collections.Generic;

namespace OfficeOpenXml.Encryption
{
    internal class SensibilityLabelInfo
    {
        public string Version { get; internal set; }
        public List<DataSpaceReference> DataSpaceMap { get; internal set; }
        public List<string> DataSpaceInfo { get; internal set; }
        public DataSpacesEncryption.TransformInfoHeader Transformation { get; internal set; }
        public string LabelXml { get; internal set; }
        public List<object> SummaryInfoProperties { get; internal set; }
        public List<object> SummaryDocumentInfoProperties { get; internal set; }
    }
}