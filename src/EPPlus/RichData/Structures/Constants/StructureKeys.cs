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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.RichData.RichValues.Errors;
using OfficeOpenXml.RichData.Structures.Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.Structures.Constants
{
    internal static class StructureKeys
    {
        internal static class Errors
        {
            internal static readonly List<ExcelRichValueStructureKey> Propagated =
                [
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.PropagatedError.ErrorType, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.PropagatedError.Propagated, RichValueDataType.String)
                ];

            internal static readonly List<ExcelRichValueStructureKey> Field =
                [
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.FieldError.ErrorType, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.FieldError.Field, RichValueDataType.String)
                ];

            internal static readonly List<ExcelRichValueStructureKey> Spill =
                [
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.Spill.ColOffset, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.Spill.ErrorType, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.Spill.RwOffset, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.Spill.SubType, RichValueDataType.Integer)
                ];

            internal static readonly List<ExcelRichValueStructureKey> WithSubType =
                [
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.WithSubType.ErrorType, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.WithSubType.SubType, RichValueDataType.Integer)
                ];

            internal static readonly List<ExcelRichValueStructureKey> Busy =
                [
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.Busy.ErrorType, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.Busy.TargetValue, RichValueDataType.RichValue)
                ];
        }

        internal static class LocalImage
        {
            internal static readonly List<ExcelRichValueStructureKey> Image =
                [
                    new ExcelRichValueStructureKey(StructureKeyNames.LocalImages.Image.RelLocalImageIdentifier, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.LocalImages.Image.CalcOrigin, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.LocalImages.Image.Text, RichValueDataType.String)
                ];
        }

        internal static class WebImage
        {
            internal static readonly List<ExcelRichValueStructureKey> Image =
                [
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.WebImageIdentifier, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.Attribution, RichValueDataType.SupportingPropertyBag),
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.Text, RichValueDataType.String),
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.ComputedImage, RichValueDataType.Bool),
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.ImageSizing, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.ImageHeight, RichValueDataType.Decimal),
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.ImageWidth, RichValueDataType.Decimal),
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.CalcOrigin, RichValueDataType.Integer),
                ];
        }

        private static Dictionary<string, Dictionary<string, RichValueDataType>> _dataTypes = new Dictionary<string, Dictionary<string, RichValueDataType>>();

        private static void RegisterKeys(string structureName, List<ExcelRichValueStructureKey> keys)
        {
            if(!_dataTypes.ContainsKey(structureName))
            {
                _dataTypes[structureName] = new Dictionary<string, RichValueDataType>();
            }
            foreach(var key in keys)
            {
                if (!_dataTypes[structureName].ContainsKey(key.Name))
                {
                    _dataTypes[structureName][key.Name] = key.DataType;
                }
            }
        }

        internal static void SortKeyNames(RichDataStructureTypes st, ref List<string> keyNames)
        {
            if((st & RichDataStructureTypes.Error) == RichDataStructureTypes.Error)
            {
                var ck = default(List<ExcelRichValueStructureKey>);
                if((st & RichDataStructureTypes.ErrorSpill) == RichDataStructureTypes.ErrorSpill)
                {
                    ck = Errors.Spill;
                }
                else if((st & RichDataStructureTypes.ErrorPropagated) == RichDataStructureTypes.ErrorPropagated)
                {
                    ck = Errors.Propagated;
                }
                else if((st & RichDataStructureTypes.ErrorWithSubType) == RichDataStructureTypes.ErrorWithSubType)
                {
                    ck = Errors.WithSubType;
                }
                else if((st & RichDataStructureTypes.ErrorField) == RichDataStructureTypes.ErrorField)
                {
                    ck = Errors.Field;
                }
                if(ck != null)
                {
                    var sortedKeys = ck.Select(x => x.Name).ToList();
                    keyNames.Sort((a, b) => sortedKeys.IndexOf(a).CompareTo(sortedKeys.IndexOf(b)));
                }
            }
        }

        internal static RichValueDataType? GetKeyDataType(string structureName, string keyName)
        {
            if(_dataTypes.Count == 0)
            {
                RegisterKeys(StructureTypes.Error, Errors.Propagated);
                RegisterKeys(StructureTypes.Error, Errors.Field);
                RegisterKeys(StructureTypes.Error, Errors.Spill);
                RegisterKeys(StructureTypes.Error, Errors.WithSubType);
                RegisterKeys(StructureTypes.LocalImage, LocalImage.Image);
                RegisterKeys(StructureTypes.WebImage, WebImage.Image);
            }
            if(_dataTypes.ContainsKey(structureName) && _dataTypes[structureName].ContainsKey(keyName))
            {
                return _dataTypes[structureName][keyName];
            }
            return null;
        }

        internal static List<ExcelRichValueStructureKey> GetDefaultKeysByType(RichDataStructureTypes structureType)
        {
            if((structureType & RichDataStructureTypes.Error) == RichDataStructureTypes.Error)
            {
                if((structureType & RichDataStructureTypes.ErrorSpill) == RichDataStructureTypes.ErrorSpill)
                {
                    return Errors.Spill;
                }
                else if ((structureType & RichDataStructureTypes.ErrorPropagated) == RichDataStructureTypes.ErrorPropagated)
                {
                    return Errors.Propagated;
                }
                else if((structureType & RichDataStructureTypes.ErrorField) == RichDataStructureTypes.ErrorField)
                {
                    return Errors.Field;
                }
                else if ((structureType & RichDataStructureTypes.ErrorWithSubType) == RichDataStructureTypes.ErrorWithSubType)
                {
                    return Errors.WithSubType;
                }
                return null;
            }
            switch (structureType)
            {
                case RichDataStructureTypes.LocalImage:
                    return LocalImage.Image;
                case RichDataStructureTypes.WebImage:
                    return WebImage.Image;
                default:
                    return null;

            }
            return null;
        }

        internal static ExcelRichValueStructureKey GetKey(RichDataStructureTypes structureType, string name)
        {
            var keys = GetDefaultKeysByType(structureType);
            return keys.FirstOrDefault(x => x.Name == name);
        }
    }
}
