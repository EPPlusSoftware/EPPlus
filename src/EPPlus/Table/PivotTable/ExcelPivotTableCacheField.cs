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
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotTableCacheField
    {
        [Flags]
        private enum DataTypeFlags
        {
            Empty = 0x1,
            String = 0x2,
            Int = 0x4,
            Number = 0x8,
            DateTime = 0x10,
            Boolean = 0x20,
            Error = 0x30
        }
        public ExcelPivotTableCacheField()
        {

        }
        public string Name
        {
            get;
            internal set;
        }
        public List<object> SharedItems
        {
            get;
            set;
        } = new List<object>();
        public ExcelPivotTableFieldNumericGroup NumericGroup { get; set; }
        public ExcelPivotTableFieldDateGroup DateGroup { get; set; }

        internal void WriteSharedItems(XmlElement fieldNode, XmlNamespaceManager nsm)
        {
            var shNode = (XmlElement)fieldNode.SelectSingleNode("d:sharedItems", nsm);

            var flags = GetFlags();

            if (flags == DataTypeFlags.String) //Strings only 
            {
                AppendSharedItems(shNode);
            }
            else if (!HasOneValueOnly(flags) && flags!=(DataTypeFlags.Int| DataTypeFlags.Number) && SharedItems.Count>1)
            {
                shNode.SetAttribute("containsMixedTypes", "1");
                if ((flags & DataTypeFlags.Empty) == DataTypeFlags.Empty)
                {
                    shNode.SetAttribute("containsBlank", "1");
                }
                if ((flags & DataTypeFlags.DateTime) == DataTypeFlags.Empty)
                {
                    shNode.SetAttribute("containsDate", "1");
                    shNode.SetAttribute("containsNonDate", "1");
                }
            }
            else
            {
                SetFlags(shNode, flags);
            }
        }

        private void SetFlags(XmlElement shNode, DataTypeFlags flags)
        {
            if((flags & DataTypeFlags.DateTime) == DataTypeFlags.DateTime)
            {
                shNode.SetAttribute("containsDate", "1");
            }
            if ((flags & DataTypeFlags.Number) == DataTypeFlags.Number)
            {
                shNode.SetAttribute("containsNumber", "1");
            }
            if ((flags & DataTypeFlags.Int) == DataTypeFlags.Int)
            {
                shNode.SetAttribute("containsInteger", "1");
            }
            if ((flags & DataTypeFlags.Empty) == DataTypeFlags.Empty)
            {
                shNode.SetAttribute("containsBlank", "1");
            }
        }

        private bool HasOneValueOnly(DataTypeFlags flags)
        {
            foreach(DataTypeFlags v in Enum.GetValues(typeof(DataTypeFlags)))
            {
                if(flags==v)
                {
                    return true;
                }
            }
            return false;
        }

        private void AppendSharedItems(XmlElement shNode)
        {
            shNode.RemoveAll();
            foreach (var si in SharedItems)
            {
                if (si == null)
                {
                    AppendItem(shNode, "m", null);
                }
                else
                {
                    var t = si.GetType();
                    switch (Type.GetTypeCode(t))
                    {
                        case TypeCode.String:
                        case TypeCode.Char:
                            AppendItem(shNode, "s", si.ToString());
                            break;
                        case TypeCode.Byte:
                        case TypeCode.SByte:
                        case TypeCode.UInt16:
                        case TypeCode.UInt32:
                        case TypeCode.UInt64:
                        case TypeCode.Int16:
                        case TypeCode.Int32:
                        case TypeCode.Int64:
                        case TypeCode.Decimal:
                        case TypeCode.Double:
                        case TypeCode.Single:
                            AppendItem(shNode, "n", ConvertUtil.GetValueForXml(si, false));
                            break;
                        case TypeCode.DateTime:
                            AppendItem(shNode, "d", ConvertUtil.GetValueForXml(si, false));
                            break;
                        case TypeCode.Boolean:
                            AppendItem(shNode, "b", ConvertUtil.GetValueForXml(si, false));
                            break;
                        case TypeCode.Empty:
                            AppendItem(shNode, "m", null);
                            break;
                        default:
                            if (t == typeof(TimeSpan))
                            {
                                AppendItem(shNode, "d", ConvertUtil.GetValueForXml(si, false));
                            }
                            else if (t == typeof(ExcelErrorValue))
                            {
                                AppendItem(shNode, "e", si.ToString());
                            }
                            else
                            {
                                AppendItem(shNode, "s", si.ToString());
                            }
                            break;
                    }

                }
            }
        }

        private DataTypeFlags GetFlags()
        {
            DataTypeFlags flags = 0;
            foreach (var si in SharedItems)
            {
                if (si == null)
                {
                    flags |= DataTypeFlags.Empty;
                }
                else
                {
                    var t = si.GetType();
                    switch (Type.GetTypeCode(t))
                    {
                        case TypeCode.String:
                        case TypeCode.Char:
                            flags |= DataTypeFlags.String;
                            break;
                        case TypeCode.Byte:
                        case TypeCode.SByte:
                        case TypeCode.UInt16:
                        case TypeCode.UInt32:
                        case TypeCode.UInt64:
                        case TypeCode.Int16:
                        case TypeCode.Int32:
                        case TypeCode.Int64:
                            flags |= (DataTypeFlags.Number|DataTypeFlags.Int);
                            break;
                        case TypeCode.Decimal:
                        case TypeCode.Double:
                        case TypeCode.Single:
                            flags |= DataTypeFlags.Number;
                            break;
                        case TypeCode.DateTime:
                            flags |= DataTypeFlags.DateTime;
                            break;
                        case TypeCode.Boolean:
                            flags |= DataTypeFlags.Boolean;
                            break;
                        case TypeCode.Empty:
                            flags |= DataTypeFlags.Empty;
                            break;
                        default:
                            if (t == typeof(TimeSpan))
                            {
                                flags |= DataTypeFlags.DateTime;
                            }
                            else if(t==typeof(ExcelErrorValue))
                            {
                                flags |= DataTypeFlags.Error;
                            }
                            else
                            {
                                flags |= DataTypeFlags.String;
                            }
                            break;
                    }
                }
            }
            return flags;
        }
        private void AppendItem(XmlElement shNode, string elementName, string value)
        {
            var e = shNode.OwnerDocument.CreateElement(elementName, ExcelPackage.schemaMain);
            if (value != null)
            {
                e.SetAttribute("v", value);
            }
            shNode.AppendChild(e);
        }
    }
}