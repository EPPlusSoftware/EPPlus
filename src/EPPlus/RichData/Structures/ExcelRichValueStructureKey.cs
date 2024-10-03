using System;

namespace OfficeOpenXml.RichData.Structures
{
    internal class ExcelRichValueStructureKey
    {
        internal ExcelRichValueStructureKey(string name, string dt)
        {
            Name = name;
            DataType = GetDataType(dt);
        }

        internal ExcelRichValueStructureKey(string name, RichValueDataType dt)
        {
            Name = name;
            DataType = dt;
        }

        private RichValueDataType GetDataType(string dt)
        {
            switch (dt)
            {
                case "spb":
                    return RichValueDataType.SupportingPropertyBag;
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
    }
}