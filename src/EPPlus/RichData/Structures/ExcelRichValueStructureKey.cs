using OfficeOpenXml.RichData.Structures.Constants;
using System;
using System.Diagnostics;

namespace OfficeOpenXml.RichData.Structures
{
    [DebuggerDisplay("Name: {Name}, Type: {DataType}")]
    internal class ExcelRichValueStructureKey
    {
        internal ExcelRichValueStructureKey(string name, string dt)
        {
            Name = name;
            DataType = GetDataType(dt);
            CheckRelation();
        }

        internal ExcelRichValueStructureKey(string name, RichValueDataType dt)
        {
            Name = name;
            DataType = dt;
            CheckRelation();
        }

        private void CheckRelation()
        {
            if (!string.IsNullOrEmpty(Name) && Name.StartsWith($"{SpecialKeyNames.Prefixes.RvRel}:"))
            {
                IsRelation = true;
                RelationName = Name.Split(':')[1];
            }
        }

        private RichValueDataType GetDataType(string dt)
        {
            switch (dt)
            {
                case "spb":
                    return RichValueDataType.SupportingPropertyBag;
                case "spba":
                    return RichValueDataType.SupportingPropertyBagArray;
                case "i":
                    return RichValueDataType.Integer;
                case "b":
                    return RichValueDataType.Bool;
                case "e":
                    return RichValueDataType.Error;
                case "s":
                    return RichValueDataType.String;
                case "r":
                    return RichValueDataType.RichValue;
                case "a":
                    return RichValueDataType.Array;
                default:
                    return RichValueDataType.Decimal;
            }
        }
        internal string GetDataTypeString()
        {
            switch (DataType)
            {
                case RichValueDataType.SupportingPropertyBag:
                    return "spb";
                case RichValueDataType.SupportingPropertyBagArray:
                    return "spba";
                case RichValueDataType.Integer:
                    return "i";
                case RichValueDataType.Bool:
                    return "b";
                case RichValueDataType.Error:
                    return "e";
                case RichValueDataType.String:
                    return "s";
                case RichValueDataType.RichValue:
                    return "r";
                case RichValueDataType.Array:
                    return "a";
                default:
                    return "d";
            }
        }

        public string Name { get; set; }
        public RichValueDataType DataType { get; set; }

        public bool IsRelation
        {
            get; private set;
        }

        public string RelationName
        {
            get; private set;
        }
    }
}