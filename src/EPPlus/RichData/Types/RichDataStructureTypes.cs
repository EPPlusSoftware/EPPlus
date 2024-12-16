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
using System;

namespace OfficeOpenXml.RichData
{
    [Flags]
    internal enum RichDataStructureTypes
    {
        None = 0,
        Error = 0x1,
        ErrorWithSubType = 0x2,
        ErrorSpill = 0x4,
        ErrorPropagated = 0x8,
        ErrorField = 0x10,
        ErrorBusy = 0x20,
        LocalImage = 0x40,
        WebImage = 0x80,
        ImageUrl = 0x100,
        LinkedEntity = 0x200,
        LinkedEntityCore = 0x400,
        LinkedEntity2 = 0x800,
        LinkedEntity2Core = 0x1000,
        FormattedNumber = 0x2000,
        Hyperlink = 0x4000,
        Entity = 0x8000,
        Array = 0x10000,
        StockHistoryCache = 0x20000,
        ExternalCodeServiceObject = 0x40000,
        SourceAttribution = 0x80000,
        Preserve = 0x100000
    }
}