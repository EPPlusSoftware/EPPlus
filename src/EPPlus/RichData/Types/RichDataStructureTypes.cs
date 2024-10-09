/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
namespace OfficeOpenXml.RichData
{
    internal enum RichDataStructureTypes
    {
        ErrorWithSubType,
        ErrorSpill,
        ErrorPropagated,
        ErrorField,
        LocalImage,
        LocalImageWithAltText,
        WebImage,
        LinkedEntity,
        LinkedEntityCore,
        LinkedEntity2,
        LinkedEntity2Core,
        FormattedNumber,
        Hyperlink,
        Array,
        Preserve
    }
}