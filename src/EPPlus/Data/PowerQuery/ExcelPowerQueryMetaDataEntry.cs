/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;

namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// An entity in the meta data entity collection for Power Query.
    /// </summary>
    public class ExcelPowerQueryMetaDataEntry
    {
        /// <summary>
        /// The valid entries and there data types. Data Types can be b=bool, s=string, i=integer, d=date/time and f=double.
        /// Entry types are case-sensitive.
        /// </summary>
        public readonly static Dictionary<string, string> ValidEntries = new Dictionary<string, string>(StringComparer.InvariantCulture)
        {
            {"AddedToDataModel", "b"},
            {"BufferNextRefresh", "b"},
            {"FillCount", "i"},
            {"FillEnabled", "b"},
            {"FillErrorCode", "s"},
            {"FillErrorCount", "i"},
            {"FillErrorMessage", "s"},
            {"FillLastUpdated", "d"},
            {"FillObjectType", "s"},
            {"FillColumnTypes", "s"},
            {"FillColumnNames", "s"},
            {"FilledCompleteResultToWorksheet", "b"},
            {"FillStatus", "s"},
            {"FillTargetNameCustomized", "b"},
            {"FillTarget", "s"},
            {"FillToDataModelEnabled", "b"},
            {"IsFunctionQuery", "b"},
            {"IsPrivate", "b"},
            {"NameUpdatedAfterFill", "b"},
            {"PublishedPackageID", "s"},
            {"PublishedPackageLastModifiedAt", "d"},
            {"QueryGroupID", "s"},
            {"QueryID", "s"},
            {"RecoveryTargetColumn", "i"},
            {"RecoveryTargetRow", "i"},
            {"RecoveryTargetSheet", "s"},
            {"RelationshipInfoContainer", "s"},
            {"ResultType", "s"},
            {"IsRelationshipDetectionEnabled", "b"},
            {"QueryGroups", "s"},
            {"Relationships", "s"},
            {"Number Of QueryGroups", "i"},
            {"QueryGroup Array", "s"},
            {"Version", "i"},
            {"Id", "s"},
            {"Name", "s"},
            {"Description", "s"},
            {"HasParent", "b"},
            {"Parent", "s"},
            {"Order", "i"}
        };
        /// <summary>
        /// Create a new <see cref="ExcelPowerQueryMetaDataEntry"/> validating the type and value according to the <see cref="ExcelPowerQueryMetaDataEntry.ValidEntries"/>
        /// </summary>
        /// <param name="type">The Entry type. </param>
        /// <param name="value">The value of the Entry. This value must be of type, string, integer, bool, datetime or double.</param>
        /// <param name="ignoreUnknownEntries">If true(default), EPPlus will ignore and preserve values not present in the <see cref="ValidEntries"/> dictionary</param>
        /// <param name="validateDataTypeForKnownEntries"></param>
        /// <exception cref="ArgumentException"></exception>
        public ExcelPowerQueryMetaDataEntry(string type, object value, bool ignoreUnknownEntries = true, bool validateDataTypeForKnownEntries = true)
        {
            if (ValidEntries.TryGetValue(type, out string dt))
            {
                if (validateDataTypeForKnownEntries)
                {
                    try
                    {
                        switch (dt)
                        {
                            case "s":
                                Value = value.ToString();
                                break;
                            case "i":
                                Value = (int)value;
                                break;
                            case "b":
                                Value = ConvertUtil.GetValueBool(value).Value;
                                break;
                            case "d":
                                Value = ConvertUtil.GetValueDate(value).Value;
                                break;
                            case "f":
                                Value = ConvertUtil.GetValueDouble(value);
                                break;
                            default:
                                throw new ArgumentException($"The data type \"{dt}\" is not valid.");
                        }
                    }
                    catch (Exception exception)
                    {
                        throw new ArgumentException($"The data type of argument {nameof(type)} must be of type \"{dt}\" (see ExcelPowerQueryMetaDataEntry.ValidEntries). value is {value}.", exception);
                    }
                }
                else
                {
                    Value = value;
                }
            }
            else
            {
                if (ignoreUnknownEntries)
                {
                    Value = value;
                }
                else
                {
                    throw new ArgumentException($"The argument {nameof(type)} must be one of the items defined in [MS-QDEFF] section 2.5.1. Please see ExcelPowerQueryMetaDataEntry.ValidEntries for valid entries and their data types.");
                }
            }
            EntryType = type;
        }
        /// <summary>
        /// The entry type. Please see <see cref="ExcelPowerQueryMetaDataEntry.ValidEntries"/> for valid entries.
        /// </summary>
        public string EntryType { get; }
        /// <summary>
        /// The value for the entry. 
        /// </summary>
        public object Value { get; }
        /// <summary>
        /// Get the value formatted according to the supplied culture.
        /// </summary>
        /// <param name="culture"></param>
        /// <returns></returns>
        public string GetValueAsText(CultureInfo culture)
        {
            culture = culture ?? Thread.CurrentThread.CurrentCulture;
            if (Value == null) return "s";
            if (Value is string s)
            {
                return "s" + s;
            }
            if (Value is int i)
            {
                return "l" + i.ToString(CultureInfo.InvariantCulture);
            }
            if (Value is DateTime dt)
            {
                return "d" + dt.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ", CultureInfo.InvariantCulture);
            }
            if (Value is bool b)
            {
                return "b" + (b ? "1" : "0");
            }
            if (Value is double d)
            {
                return "f" + d.ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                return "s" + Value.ToString(); //You should never end up here.
            }
        }
        /// <summary>
        /// Return the object as a string.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return EntryType + ":" + GetValueAsText(null);
        }
    }
}