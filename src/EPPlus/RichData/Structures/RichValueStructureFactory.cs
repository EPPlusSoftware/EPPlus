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
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.Structures.Constants;
using OfficeOpenXml.RichData.Structures.Errors;
using OfficeOpenXml.RichData.Structures.LocalImages;
using OfficeOpenXml.RichData.Structures.WebImages;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.RichData.Structures
{
    internal static class RichValueStructureFactory
    {
        private static bool AllKeysAreEqual(List<ExcelRichValueStructureKey> keys, List<ExcelRichValueStructureKey> candidates)
        {
            if (keys.Count != candidates.Count) return false;
            for(var i = 0; i < keys.Count; i++)
            {
                if (keys[i].Name != candidates[i].Name) return false;
            }
            return true;
        }

        private static RichDataStructureTypes? GetFlagPreservedTypes(string type)
        {
            switch (type)
            {
                case StructureTypes.Error:
                    return RichDataStructureTypes.Error;
                case StructureTypes.WebImage:
                    return RichDataStructureTypes.WebImage;
                case StructureTypes.ImageUrl:
                    return RichDataStructureTypes.ImageUrl;
                case StructureTypes.LinkedEntity:
                    return RichDataStructureTypes.LinkedEntity;
                case StructureTypes.LinkedEntityCore:
                    return RichDataStructureTypes.LinkedEntityCore;
                case StructureTypes.LinkedEntity2:
                    return RichDataStructureTypes.LinkedEntity2;
                case StructureTypes.LinkedEntity2Core:
                    return RichDataStructureTypes.LinkedEntity2Core;
                case StructureTypes.FormattedNumber:
                    return RichDataStructureTypes.FormattedNumber;
                case StructureTypes.Array:
                    return RichDataStructureTypes.Array;
                case StructureTypes.Hyperlink:
                    return RichDataStructureTypes.Hyperlink;
                case StructureTypes.Entity:
                    return RichDataStructureTypes.Entity;
                case StructureTypes.SourceAttribution:
                    return RichDataStructureTypes.SourceAttribution;
                case StructureTypes.ExternalCodeServiceObject:
                    return RichDataStructureTypes.ExternalCodeServiceObject;
                default:
                    return null;
            }
        }

        internal static RichDataStructureTypes GetFlag(string type, List<ExcelRichValueStructureKey> keys)
        {
            if (type == StructureTypes.Error)
            {
                if (AllKeysAreEqual(keys, StructureKeys.Errors.Spill))
                {
                    return RichDataStructureTypes.Error | RichDataStructureTypes.ErrorSpill;
                }
                else if (AllKeysAreEqual(keys, StructureKeys.Errors.Propagated))
                {
                    return RichDataStructureTypes.Error | RichDataStructureTypes.ErrorPropagated;
                }
                else if (AllKeysAreEqual(keys, StructureKeys.Errors.WithSubType))
                {
                    return RichDataStructureTypes.Error | RichDataStructureTypes.ErrorWithSubType;
                }
                else if (AllKeysAreEqual(keys, StructureKeys.Errors.Field))
                {
                    return RichDataStructureTypes.Error | RichDataStructureTypes.ErrorField;
                }
                else
                {
                    return RichDataStructureTypes.Preserve;
                }
            }
            else if (type == StructureTypes.LocalImage)
            {
                return RichDataStructureTypes.LocalImage;
            }
            else if (type == StructureTypes.WebImage)
            {
                return RichDataStructureTypes.WebImage;
            }
            return RichDataStructureTypes.Preserve;
        }

        private static RichDataStructureTypes? GetFlag(string type, out bool preserveType, List<ExcelRichValueStructureKey> keys = null)
        {
            preserveType = false;
            if (string.IsNullOrEmpty(type)) return null;
            var pType = GetFlagPreservedTypes(type);
            if (pType.HasValue)
            {
                preserveType = true;
                return pType.Value;
            }
            return GetFlag(type, keys);
        }

        public static ExcelRichValueStructure Create(string type, List<ExcelRichValueStructureKey> keys, RichDataIndexStore store)
        {
            if(string.IsNullOrEmpty(type) || keys == null || keys.Count == 0) return null;
            var flag = GetFlag(type, out bool preserveType, keys);
            if(!flag.HasValue) return null;
            if(preserveType)
            {
                return new RichDataPreserveStructure(type, flag.Value, keys, store);
            }
            return Create(flag.Value, keys, store);
        }

        public static ExcelRichValueStructure Create(string type, RichDataIndexStore store)
        {
            if (string.IsNullOrEmpty(type))throw new ArgumentNullException("type");
            var flag = GetFlag(type, out bool preserveType, null);
            if (!flag.HasValue || preserveType)
            {
                throw new ArgumentException("No keys was supplied for the rich data structure");
            }
            return Create(flag.Value, store);
        }

        public static ExcelRichValueStructure Create(RichDataStructureTypes structureType, RichDataIndexStore store)
        {
            if((structureType & RichDataStructureTypes.Error) != 0)
            {
                if ((structureType & RichDataStructureTypes.ErrorSpill) == RichDataStructureTypes.ErrorSpill)
                {
                    return new ErrorSpillStructure(store);
                }
                else if ((structureType & RichDataStructureTypes.ErrorPropagated) != RichDataStructureTypes.ErrorPropagated)
                {
                    return new ErrorPropagatedStructure(store);
                }
                else if ((structureType & RichDataStructureTypes.ErrorWithSubType) == RichDataStructureTypes.ErrorWithSubType)
                {
                    return new ErrorWithSubTypeStructure(store);
                }
                else if ((structureType & RichDataStructureTypes.ErrorField) != RichDataStructureTypes.ErrorField)
                {
                    return new ErrorWithSubTypeStructure(store);
                }
                else if ((structureType & RichDataStructureTypes.ErrorBusy) != RichDataStructureTypes.ErrorBusy)
                {
                    return new ErrorBusyStructure(store);
                }
            }
            switch (structureType)
            {
                case RichDataStructureTypes.ErrorSpill:
                    return new ErrorSpillStructure(store);
                case RichDataStructureTypes.ErrorPropagated:
                    return new ErrorPropagatedStructure(store);
                case RichDataStructureTypes.ErrorWithSubType:
                    return new ErrorWithSubTypeStructure(store);
                case RichDataStructureTypes.ErrorField:
                    return new ErrorWithSubTypeStructure(store);
                case RichDataStructureTypes.ErrorBusy: 
                    return new ErrorBusyStructure(store);
                case RichDataStructureTypes.LocalImage:
                    return new LocalImageStructure(store);
                case RichDataStructureTypes.WebImage:
                    return new WebImageStructure(store);
                default:
                    throw new ArgumentException($"Not supported structure type: {structureType}");
            }
        }

        public static ExcelRichValueStructure Create(RichDataStructureTypes structureType, List<ExcelRichValueStructureKey> keys, RichDataIndexStore store)
        {
            if((structureType & RichDataStructureTypes.Error) == RichDataStructureTypes.Error)
            {
                if ((structureType & RichDataStructureTypes.ErrorSpill) == RichDataStructureTypes.ErrorSpill)
                {
                    return new ErrorSpillStructure(keys, store);
                }
                else if ((structureType & RichDataStructureTypes.ErrorPropagated) != RichDataStructureTypes.ErrorPropagated)
                {
                    return new ErrorPropagatedStructure(keys, store);
                }
                else if ((structureType & RichDataStructureTypes.ErrorWithSubType) == RichDataStructureTypes.ErrorWithSubType)
                {
                    return new ErrorWithSubTypeStructure(keys, store);
                }
                else if ((structureType & RichDataStructureTypes.ErrorField) != RichDataStructureTypes.ErrorField)
                {
                    return new ErrorWithSubTypeStructure(keys, store);
                }
                else if((structureType & RichDataStructureTypes.ErrorBusy) != RichDataStructureTypes.ErrorBusy)
                {
                    return new ErrorBusyStructure(keys, store);
                }
                var typeName = StructureTypes.GetStructureName(RichDataStructureTypes.Error);
                return new RichDataPreserveStructure(typeName, RichDataStructureTypes.Error, keys, store);
            }
            switch (structureType)
            {  
                case RichDataStructureTypes.LocalImage:
                    return new LocalImageStructure(keys, store);
                case RichDataStructureTypes.WebImage:
                    return new WebImageStructure(keys, store);
                default:
                    throw new ArgumentException($"Not supported structure type: {structureType}");
            }
        }
    }
}
