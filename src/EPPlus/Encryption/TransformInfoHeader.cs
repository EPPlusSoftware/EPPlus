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
namespace OfficeOpenXml.Encryption
{
    internal class TransformInfoHeader
    {
        public int TransformType { get; set; }
        public string TransformId { get; set; }
        public string TransformName { get; set; }
        public string ReaderVersion { get; set; }
        public string UpdaterVersion { get; set; }
        public string WriterVersion { get; set; }
        public string LicenseXrML { get; set; }
    }
}
