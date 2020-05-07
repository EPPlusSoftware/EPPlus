using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    internal static class TemperatureConverter
    {
        private static double Cel2Fah(double c)
        {
            return (c * 9) / 5 + 32;
        }

        private static double Fah2Cel(double f)
        {
            return (f - 32) * 5 / 9;
        }

        private static double Cel2Kel(double c)
        {
            return c + 273.15;
        }

        private static double Kel2Cel(double k)
        {
            return k - 273.15;
        }

        private static double Fah2Kel(double f)
        {
            var c = Fah2Cel(f);
            return Cel2Kel(c);
        }

        private static double Kel2Fah(double k)
        {
            var c = Kel2Cel(k);
            return Cel2Fah(c);
        }

        public static bool IsValidUnit(string candidate)
        {
            return new List<string>
            {
                "C",
                "cel",
                "F",
                "fah",
                "K",
                "kel"
            }.Contains(candidate);
        }


        public static Dictionary<string, Func<double, double>> Conversions = new Dictionary<string, Func<double, double>>
        {
            { "C>F", Cel2Fah },
            { "C>fah", Cel2Fah },
            { "cel>F", Cel2Fah },
            { "cel>fah", Cel2Fah },
            { "F>C", Fah2Cel },
            { "fah>C", Fah2Cel },
            { "F>cel", Fah2Cel },
            { "fah>cel", Fah2Cel },
            { "C>K", Cel2Kel },
            { "C>kel", Cel2Kel },
            { "cel>kel", Cel2Kel },
            { "cel>K", Cel2Kel },
            { "F>K", Fah2Kel },
            { "fah>K", Fah2Kel },
            { "F>kel", Fah2Kel },
            { "fah>kel", Fah2Kel },
            { "K>C", Kel2Cel },
            { "K->cel", Kel2Cel },
            { "kel->C", Kel2Cel },
            { "kel->cel", Kel2Cel },
            { "K>F", Kel2Fah },
            { "K>fah", Kel2Fah },
            { "kel>F", Kel2Fah },
            { "kel>fah", Kel2Fah }
        };

       

        public static bool IsTempMapping(string from, string to)
        {
            return Conversions.ContainsKey($"{from}>{to}");
        }
    }
}
