using System;

namespace OfficeOpenXml.RichData
{
    internal class ExcelRichValueStructureKey
    {
        internal ExcelRichValueStructureKey(string name, string dt)
        {
            Name = name;
            DataType = GetDataType(dt);
        }

        private RichValueDataType GetDataType(string dt)
        {
            switch(dt)
            {
                case "d":
                    return RichValueDataType.Decimal;
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
                    return RichValueDataType.SupportingPropertyBag;
            }
        }
        internal string GetDataTypeString()
        {
            switch (DataType)
            {
                case RichValueDataType.Decimal:
                    return "d";
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
                    return "spb";
            }
        }

        public string Name { get; set; }
        public RichValueDataType DataType { get; set; }
    }
}