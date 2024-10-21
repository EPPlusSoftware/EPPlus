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
using OfficeOpenXml.Encryption;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.Structures.Constants;
using OfficeOpenXml.RichData.Structures.Errors;
using OfficeOpenXml.RichData.Structures.LocalImages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

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
            if (type == StructureTypes.Error)
            {
                if (AllKeysAreEqual(keys, StructureKeys.Errors.Spill))
                {
                    return RichDataStructureTypes.ErrorSpill;
                }    
                else if (AllKeysAreEqual(keys, StructureKeys.Errors.Propagated))
                {
                    return RichDataStructureTypes.ErrorPropagated;
                }
                else if (AllKeysAreEqual(keys, StructureKeys.Errors.WithSubType))
                {
                    return RichDataStructureTypes.ErrorWithSubType;
                }
                else if(AllKeysAreEqual(keys, StructureKeys.Errors.Field))
                {
                    return RichDataStructureTypes.ErrorField;
                }
                else
                {
                    return RichDataStructureTypes.Preserve;
                }
            }
            else if(type == StructureTypes.LocalImage)
            {
                if(AllKeysAreEqual(keys, StructureKeys.LocalImage.Image))
                {
                    return RichDataStructureTypes.LocalImage;
                }
                else if(AllKeysAreEqual(keys, StructureKeys.LocalImage.ImageAltText))
                {
                    return RichDataStructureTypes.LocalImageWithAltText;
                }
            }
            return RichDataStructureTypes.Preserve;
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
                case RichDataStructureTypes.LocalImage:
                    return new LocalImageStructure(store);
                case RichDataStructureTypes.LocalImageWithAltText:
                    return new LocalImageWithAltTextStructure(store);
                default:
                    throw new ArgumentException($"Not supported structure type: {structureType}");
            }
        }

        public static ExcelRichValueStructure Create(RichDataStructureTypes structureType, List<ExcelRichValueStructureKey> keys, RichDataIndexStore store)
        {
            switch (structureType)
            {
                case RichDataStructureTypes.ErrorSpill:
                    return new ErrorSpillStructure(keys, store);
                case RichDataStructureTypes.ErrorPropagated:
                    return new ErrorPropagatedStructure(keys, store);
                case RichDataStructureTypes.ErrorWithSubType:
                    return new ErrorWithSubTypeStructure(keys, store);
                case RichDataStructureTypes.ErrorField:
                    return new ErrorWithSubTypeStructure(keys, store);
                case RichDataStructureTypes.LocalImage:
                    return new LocalImageStructure(keys, store);
                case RichDataStructureTypes.LocalImageWithAltText:
                    return new LocalImageWithAltTextStructure(keys, store);
                default:
                    throw new ArgumentException($"Not supported structure type: {structureType}");
            }
        }
    }
}
