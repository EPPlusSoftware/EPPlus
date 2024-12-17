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

namespace OfficeOpenXml.RichData.Structures.Constants
{
    internal static class StructureTypes
    {
        public const string Error = "_error";
        public const string LocalImage = "_localImage";
        public const string WebImage = "_webimage";
        public const string ImageUrl = "_imageurl";
        public const string LinkedEntity = "_linkedentity";
        public const string LinkedEntity2 = "_linkedentity2";
        public const string LinkedEntityCore = "_linkedentitycore";
        public const string LinkedEntity2Core = "_linkedentity2core";
        public const string FormattedNumber = "_formattednumber";
        public const string Hyperlink = "_hyperlink";
        public const string Array = "_array";
        public const string Entity = "_entity";
        public const string StockHistoryCache = "_stockhistorycache";
        public const string ExternalCodeServiceObject = "_python";
        public const string SourceAttribution = "_sourceattribution";

        internal static RichDataStructureTypes GetStructureType(string name)
        {
            switch(name)
            {
                case Error:
                    return RichDataStructureTypes.Error;
                case LocalImage:
                    return RichDataStructureTypes.LocalImage;
                case WebImage:
                    return RichDataStructureTypes.WebImage;
                case LinkedEntity:
                    return RichDataStructureTypes.LinkedEntity;
                case LinkedEntity2:
                    return RichDataStructureTypes.LinkedEntity2;
                case LinkedEntityCore:
                    return RichDataStructureTypes.LinkedEntityCore;
                case LinkedEntity2Core:
                    return RichDataStructureTypes.LinkedEntity2Core;
                case FormattedNumber:
                    return RichDataStructureTypes.FormattedNumber;
                case Hyperlink:
                    return RichDataStructureTypes.Hyperlink;
                case Array:
                    return RichDataStructureTypes.Array;
                case Entity:
                    return RichDataStructureTypes.Entity;
                case StockHistoryCache:
                    return RichDataStructureTypes.StockHistoryCache;
                case ExternalCodeServiceObject:
                    return RichDataStructureTypes.ExternalCodeServiceObject;
                case SourceAttribution:
                    return RichDataStructureTypes.SourceAttribution;
                default:
                    throw new InvalidOperationException("Invalid structure type: " + name);
                    
            }
        }

        internal static string GetStructureName(RichDataStructureTypes structureType)
        {
            switch(structureType)
            {
                case RichDataStructureTypes.Error:
                    return Error;
                case RichDataStructureTypes.LocalImage:
                    return LocalImage;
                case RichDataStructureTypes.WebImage:
                    return WebImage;
                case RichDataStructureTypes.LinkedEntity:
                    return LinkedEntity;
                case RichDataStructureTypes.LinkedEntity2:
                    return LinkedEntity2;
                case RichDataStructureTypes.LinkedEntityCore:
                    return LinkedEntityCore;
                case RichDataStructureTypes.LinkedEntity2Core:
                    return LinkedEntity2Core;
                case RichDataStructureTypes.FormattedNumber:
                    return FormattedNumber;
                case RichDataStructureTypes.Hyperlink:
                    return Hyperlink;
                case RichDataStructureTypes.Array:
                    return Array;
                case RichDataStructureTypes.Entity:
                    return Entity;
                case RichDataStructureTypes.StockHistoryCache:
                    return StockHistoryCache;
                case RichDataStructureTypes.ExternalCodeServiceObject:
                    return ExternalCodeServiceObject;
                case RichDataStructureTypes.SourceAttribution:
                    return SourceAttribution;
                default:
                    throw new NotImplementedException("Not supported structureType: " + structureType.ToString());
            }
        }
    }
}
