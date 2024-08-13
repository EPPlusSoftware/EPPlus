/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using MathObj = System.Math;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    /// <summary>
    /// This static class contains all the setup, definitions and methods needed for Excel's Convert function
    /// </summary>
    internal class Conversions
    {
        #region Types
        /// <summary>
        /// Types of mapping groups
        /// </summary>
        private enum UnitTypes
        {
            Distance,
            Time,
            WeightAndMass,
            Speed,
            Area,
            Liquid,
            Power,
            Magnetism,
            Force,
            Pressure,
            Energy,
            InformationUnits
        }

        /// <summary>
        /// A mapping definition
        /// </summary>
        private struct Unit
        {
            public Unit(double value, UnitTypes type)
            {
                Value = value;
                UnitType = type;
            }

            public double Value { get; private set; }

            public UnitTypes UnitType { get; private set; }
        }

        /// <summary>
        /// Represents a prefix and its value, such as the k in km (kilo-meters).
        /// </summary>
        private struct Prefix
        {
            public Prefix(string abbrevation, double value)
            {
                Abbrevation = abbrevation;
                Value = value;
            }

            public string Abbrevation { get; private set; }

            public double Value { get; private set; }
        }

        #endregion

        private static Dictionary<string, Unit> _conversions = new Dictionary<string, Unit>();
        private static List<Prefix> _metricPrefixes = null;
        private static List<Prefix> _binaryPrefixes = null;
        private static bool _initialized = false;
        private static Lock _syncRoot = new Lock();

        #region Initialization methods

        private static void InitMetricPrefixList()
        {
            _metricPrefixes = new List<Prefix>
            {
                new Prefix("y", MathObj.Pow(10, -24)),
                new Prefix("z", MathObj.Pow(10, -21)),
                new Prefix("a", MathObj.Pow(10, -18)),
                new Prefix("f", MathObj.Pow(10, -15)),
                new Prefix("p", MathObj.Pow(10, -12)),
                new Prefix("n", MathObj.Pow(10, -9)),
                new Prefix("u", MathObj.Pow(10, -6)),
                new Prefix("m", MathObj.Pow(10, -3)),
                new Prefix("c", MathObj.Pow(10, -2)),
                new Prefix("d", MathObj.Pow(10, -1)),
                new Prefix("da", MathObj.Pow(10, 1)),
                new Prefix("e", MathObj.Pow(10, 1)),
                new Prefix("h", MathObj.Pow(10, 2)),
                new Prefix("k", MathObj.Pow(10, 3)),
                new Prefix("M", MathObj.Pow(10, 6)),
                new Prefix("G", MathObj.Pow(10, 9)),
                new Prefix("T", MathObj.Pow(10, 12)),
                new Prefix("P", MathObj.Pow(10, 15)),
                new Prefix("E", MathObj.Pow(10, 18)),
                new Prefix("Z", MathObj.Pow(10, 21)),
                new Prefix("Y", MathObj.Pow(10, 24))
            };
        }

        private static void InitBinaryPrefixList()
        {
            _binaryPrefixes = new List<Prefix>
            {
                new Prefix("ki", MathObj.Pow(2, 10)),
                new Prefix("Mi", MathObj.Pow(2, 20)),
                new Prefix("Gi", MathObj.Pow(2, 30)),
                new Prefix("Ti", MathObj.Pow(2, 40)),
                new Prefix("Pi", MathObj.Pow(2, 50)),
                new Prefix("Ei", MathObj.Pow(2, 60)),
                new Prefix("Zi", MathObj.Pow(2, 70)),
                new Prefix("Yi", MathObj.Pow(2, 80)),
            };
        }

        private static void AddMetricPrefix(string abbrevation)
        {
            if (!_conversions.ContainsKey(abbrevation)) return;
            var conversion = _conversions[abbrevation];
            foreach(var prefix in _metricPrefixes)
            {
                switch(conversion.UnitType)
                {
                    case UnitTypes.Area:
                        _conversions[prefix.Abbrevation + abbrevation] = new Unit(MathObj.Pow(prefix.Value, 2) * conversion.Value, conversion.UnitType);
                        break;
                    case UnitTypes.Liquid:
                        _conversions[prefix.Abbrevation + abbrevation] = new Unit(MathObj.Pow(prefix.Value, 3) * conversion.Value, conversion.UnitType);
                        break;
                    default:
                        _conversions[prefix.Abbrevation + abbrevation] = new Unit(prefix.Value * conversion.Value, conversion.UnitType);
                        break;
                }
            }
        }

        private static void AddBinaryPrefix(string abbrevation)
        {
            if (!_conversions.ContainsKey(abbrevation)) return;
            var conversion = _conversions[abbrevation];
            foreach(var prefix in _binaryPrefixes)
            {
                _conversions[prefix.Abbrevation + abbrevation] = new Unit(prefix.Value * conversion.Value, conversion.UnitType);
            }
        }

        private static void InitMetricPrefixes()
        {
            InitMetricPrefixList();
            AddMetricPrefix("m");
            AddMetricPrefix("ang");
            AddMetricPrefix("g");
            AddMetricPrefix("m/h");
            AddMetricPrefix("m/hr");
            AddMetricPrefix("m/s");
            AddMetricPrefix("m/sec");
            AddMetricPrefix("m2");
            AddMetricPrefix("ar");
            AddMetricPrefix("ang2");
            AddMetricPrefix("ang3");
            AddMetricPrefix("ang^3");
            AddMetricPrefix("l");
            AddMetricPrefix("lt");
            AddMetricPrefix("m3");
            AddMetricPrefix("m^3");
            AddMetricPrefix("W");
            AddMetricPrefix("w");
            AddMetricPrefix("N");
            AddMetricPrefix("dyn");
            AddMetricPrefix("J");
            AddMetricPrefix("Wh");
            AddMetricPrefix("e");
        }

        private static void InitBinaryPrefixes()
        {
            InitBinaryPrefixList();
            AddBinaryPrefix("bit");
            AddBinaryPrefix("byte");
        }

        private static void Init()
        {
            _conversions["m"] = new Unit(Distance.Meter, UnitTypes.Distance);
            _conversions["mi"] = new Unit(Distance.StatueMile, UnitTypes.Distance);
            _conversions["Nmi"] = new Unit(Distance.Nmi, UnitTypes.Distance);
            _conversions["in"] = new Unit(Distance.Inch, UnitTypes.Distance);
            _conversions["ft"] = new Unit(Distance.Foot, UnitTypes.Distance);
            _conversions["yd"] = new Unit(Distance.Yard, UnitTypes.Distance);
            _conversions["ang"] = new Unit(Distance.Angstrom, UnitTypes.Distance);
            _conversions["pica"] = new Unit(Distance.Pica, UnitTypes.Distance);
            _conversions["Pica"] = new Unit(Distance.Picapt, UnitTypes.Distance);
            _conversions["Picapt"] = new Unit(Distance.Picapt, UnitTypes.Distance);
            _conversions["ell"] = new Unit(Distance.Ell, UnitTypes.Distance);
            _conversions["ly"] = new Unit(Distance.LightYear, UnitTypes.Distance);
            _conversions["parsec"] = new Unit(Distance.Parsec, UnitTypes.Distance);
            _conversions["pc"] = new Unit(Distance.Parsec, UnitTypes.Distance);

            _conversions["yr"] = new Unit(Time.Year, UnitTypes.Time);
            _conversions["day"] = new Unit(Time.Day, UnitTypes.Time);
            _conversions["hr"] = new Unit(Time.Hour, UnitTypes.Time);
            _conversions["mn"] = new Unit(Time.Minute, UnitTypes.Time);
            _conversions["sec"] = new Unit(Time.Second, UnitTypes.Time);

            _conversions["g"] = new Unit(WeightAndMass.Grain, UnitTypes.WeightAndMass);
            _conversions["sg"] = new Unit(WeightAndMass.Slug, UnitTypes.WeightAndMass);
            _conversions["lbm"] = new Unit(WeightAndMass.PoundMass, UnitTypes.WeightAndMass);
            _conversions["u"] = new Unit(WeightAndMass.U, UnitTypes.WeightAndMass);
            _conversions["ozm"] = new Unit(WeightAndMass.OunceMass, UnitTypes.WeightAndMass);
            _conversions["grain"] = new Unit(WeightAndMass.Grain, UnitTypes.WeightAndMass);
            _conversions["cwt"] = new Unit(WeightAndMass.UsHundredweight, UnitTypes.WeightAndMass);
            _conversions["shweight"] = new Unit(WeightAndMass.UsHundredweight, UnitTypes.WeightAndMass);
            _conversions["uk_cwt"] = new Unit(WeightAndMass.ImperialHundredweight, UnitTypes.WeightAndMass);
            _conversions["lcwt"] = new Unit(WeightAndMass.ImperialHundredweight, UnitTypes.WeightAndMass);
            _conversions["hweight"] = new Unit(WeightAndMass.ImperialHundredweight, UnitTypes.WeightAndMass);
            _conversions["stone"] = new Unit(WeightAndMass.Stone, UnitTypes.WeightAndMass);
            _conversions["ton"] = new Unit(WeightAndMass.Ton, UnitTypes.WeightAndMass);
            _conversions["uk_ton"] = new Unit(WeightAndMass.ImperialTon, UnitTypes.WeightAndMass);
            _conversions["LTON"] = new Unit(WeightAndMass.ImperialTon, UnitTypes.WeightAndMass);
            _conversions["brton"] = new Unit(WeightAndMass.ImperialTon, UnitTypes.WeightAndMass);

            _conversions["admkn"] = new Unit(Speed.AdmiralKnot, UnitTypes.Speed);
            _conversions["kn"] = new Unit(Speed.Knot, UnitTypes.Speed);
            _conversions["m/h"] = new Unit(Speed.MetersPerHour, UnitTypes.Speed);
            _conversions["m/hr"] = new Unit(Speed.MetersPerHour, UnitTypes.Speed);
            _conversions["m/s"] = new Unit(Speed.MetersPerSecond, UnitTypes.Speed);
            _conversions["m/sec"] = new Unit(Speed.MetersPerSecond, UnitTypes.Speed);
            _conversions["mph"] = new Unit(Speed.MilesPerHour, UnitTypes.Speed);

            _conversions["uk_acre"] = new Unit(Area.InternationalAcre, UnitTypes.Area);
            _conversions["us_acre"] = new Unit(Area.UsServeyAcre, UnitTypes.Area);
            _conversions["ang2"] = new Unit(Area.SquareAngstrom, UnitTypes.Area);
            _conversions["ang^2"] = new Unit(Area.SquareAngstrom, UnitTypes.Area);
            _conversions["ar"] = new Unit(Area.Are, UnitTypes.Area);
            _conversions["ft2"] = new Unit(Area.SquareFeet, UnitTypes.Area);
            _conversions["ft^2"] = new Unit(Area.SquareFeet, UnitTypes.Area);
            _conversions["ha"] = new Unit(Area.Hectare, UnitTypes.Area);
            _conversions["in2"] = new Unit(Area.SquareInches, UnitTypes.Area);
            _conversions["in^2"] = new Unit(Area.SquareInches, UnitTypes.Area);
            _conversions["ly2"] = new Unit(Area.SquareLightYear, UnitTypes.Area);
            _conversions["ly^2"] = new Unit(Area.SquareLightYear, UnitTypes.Area);
            _conversions["m2"] = new Unit(Area.SquareMeter, UnitTypes.Area);
            _conversions["m^2"] = new Unit(Area.SquareMeter, UnitTypes.Area);
            _conversions["Morgen"] = new Unit(Area.Morgen, UnitTypes.Area);
            _conversions["mi2"] = new Unit(Area.SquareMiles, UnitTypes.Area);
            _conversions["mi^2"] = new Unit(Area.SquareMiles, UnitTypes.Area);
            _conversions["Nmi2"] = new Unit(Area.SquareNauticalMiles, UnitTypes.Area);
            _conversions["Nmi^2"] = new Unit(Area.SquareNauticalMiles, UnitTypes.Area);
            _conversions["Picapt2"] = new Unit(Area.SquarePica, UnitTypes.Area);
            _conversions["Picapt^2"] = new Unit(Area.SquarePica, UnitTypes.Area);
            _conversions["Pica2"] = new Unit(Area.SquarePica, UnitTypes.Area);
            _conversions["Pica^2"] = new Unit(Area.SquarePica, UnitTypes.Area);
            _conversions["yd2"] = new Unit(Area.SquareYards, UnitTypes.Area);
            _conversions["yd^2"] = new Unit(Area.SquareYards, UnitTypes.Area);

            _conversions["tsp"] = new Unit(Liquid.Teaspoon, UnitTypes.Liquid);
            _conversions["tbs"] = new Unit(Liquid.Tablespoon, UnitTypes.Liquid);
            _conversions["oz"] = new Unit(Liquid.FluidOunce, UnitTypes.Liquid);
            _conversions["cup"] = new Unit(Liquid.Cup, UnitTypes.Liquid);
            _conversions["pt"] = new Unit(Liquid.UsPint, UnitTypes.Liquid);
            _conversions["us_pt"] = new Unit(Liquid.UsPint, UnitTypes.Liquid);
            _conversions["uk_pt"] = new Unit(Liquid.UkPint, UnitTypes.Liquid);
            _conversions["qt"] = new Unit(Liquid.Quart, UnitTypes.Liquid);
            _conversions["uk_qt"] = new Unit(Liquid.ImperialQuart, UnitTypes.Liquid);
            _conversions["gal"] = new Unit(Liquid.Gallon, UnitTypes.Liquid);
            _conversions["l"] = new Unit(Liquid.Liter, UnitTypes.Liquid);
            _conversions["lt"] = new Unit(Liquid.Liter, UnitTypes.Liquid);
            _conversions["ang3"] = new Unit(Liquid.CubicAngstrom, UnitTypes.Liquid);
            _conversions["ang^3"] = new Unit(Liquid.CubicAngstrom, UnitTypes.Liquid);
            _conversions["barrel"] = new Unit(Liquid.UsOilBarrel, UnitTypes.Liquid);
            _conversions["bushel"] = new Unit(Liquid.UsBushel, UnitTypes.Liquid);
            _conversions["ft3"] = new Unit(Liquid.CubicFeet, UnitTypes.Liquid);
            _conversions["ft^3"] = new Unit(Liquid.CubicFeet, UnitTypes.Liquid);
            _conversions["in3"] = new Unit(Liquid.CubicInch, UnitTypes.Liquid);
            _conversions["in^3"] = new Unit(Liquid.CubicInch, UnitTypes.Liquid);
            _conversions["ly3"] = new Unit(Liquid.CubicLightYear, UnitTypes.Liquid);
            _conversions["ly^3"] = new Unit(Liquid.CubicLightYear, UnitTypes.Liquid);
            _conversions["m3"] = new Unit(Liquid.CubicMeter, UnitTypes.Liquid);
            _conversions["m^3"] = new Unit(Liquid.CubicMeter, UnitTypes.Liquid);
            _conversions["mi3"] = new Unit(Liquid.CubicMile, UnitTypes.Liquid);
            _conversions["mi^3"] = new Unit(Liquid.CubicMile, UnitTypes.Liquid);
            _conversions["yd3"] = new Unit(Liquid.CubicYard, UnitTypes.Liquid);
            _conversions["yd^3"] = new Unit(Liquid.CubicYard, UnitTypes.Liquid);
            _conversions["Nmi3"] = new Unit(Liquid.CubicNauticalMile, UnitTypes.Liquid);
            _conversions["Nmi^3"] = new Unit(Liquid.CubicNauticalMile, UnitTypes.Liquid);
            _conversions["Picapt3"] = new Unit(Liquid.CubicPica, UnitTypes.Liquid);
            _conversions["Picapt^3"] = new Unit(Liquid.CubicPica, UnitTypes.Liquid);
            _conversions["Pica3"] = new Unit(Liquid.CubicPica, UnitTypes.Liquid);
            _conversions["Pica^3"] = new Unit(Liquid.CubicPica, UnitTypes.Liquid);
            _conversions["GRT"] = new Unit(Liquid.GrossRegistredTon, UnitTypes.Liquid);
            _conversions["regton"] = new Unit(Liquid.GrossRegistredTon, UnitTypes.Liquid);
            _conversions["MTON"] = new Unit(Liquid.MeasurementTon, UnitTypes.Liquid);

            _conversions["HP"] = new Unit(Power.Horsepower, UnitTypes.Power);
            _conversions["h"] = new Unit(Power.Horsepower, UnitTypes.Power);
            _conversions["W"] = new Unit(Power.Watt, UnitTypes.Power);
            _conversions["w"] = new Unit(Power.Watt, UnitTypes.Power);
            _conversions["PS"] = new Unit(Power.PS, UnitTypes.Power);

            _conversions["T"] = new Unit(Magnetism.Tesla, UnitTypes.Magnetism);
            _conversions["ga"] = new Unit(Magnetism.Gauss, UnitTypes.Magnetism);

            _conversions["N"] = new Unit(Force.Newton, UnitTypes.Force);
            _conversions["dyn"] = new Unit(Force.Dyne, UnitTypes.Force);
            _conversions["dy"] = new Unit(Force.Dyne, UnitTypes.Force);
            _conversions["lbf"] = new Unit(Force.PoundForce, UnitTypes.Force);
            _conversions["pond"] = new Unit(Force.Pond, UnitTypes.Force);

            _conversions["Pa"] = new Unit(Pressure.Pascal, UnitTypes.Pressure);
            _conversions["p"] = new Unit(Pressure.Pascal, UnitTypes.Pressure);
            _conversions["atm"] = new Unit(Pressure.Atmosphere, UnitTypes.Pressure);
            _conversions["at"] = new Unit(Pressure.Atmosphere, UnitTypes.Pressure);
            _conversions["mmHg"] = new Unit(Pressure.MmOfMercury, UnitTypes.Pressure);
            _conversions["psi"] = new Unit(Pressure.Psi, UnitTypes.Pressure);
            _conversions["Torr"] = new Unit(Pressure.Torr, UnitTypes.Pressure);

            _conversions["J"] = new Unit(Energy.Joule, UnitTypes.Energy);
            _conversions["e"] = new Unit(Energy.Erg, UnitTypes.Energy);
            _conversions["c"] = new Unit(Energy.ThermodynamicCalorie, UnitTypes.Energy);
            _conversions["cal"] = new Unit(Energy.ItCalorie, UnitTypes.Energy);
            _conversions["eV"] = new Unit(Energy.ElectronVolt, UnitTypes.Energy);
            _conversions["ev"] = new Unit(Energy.ElectronVolt, UnitTypes.Energy);
            _conversions["HPh"] = new Unit(Energy.HorsePowerHour, UnitTypes.Energy);
            _conversions["Wh"] = new Unit(Energy.WattHour, UnitTypes.Energy);
            _conversions["flb"] = new Unit(Energy.FootPound, UnitTypes.Energy);
            _conversions["BTU"] = new Unit(Energy.Btu, UnitTypes.Energy);
            _conversions["btu"] = new Unit(Energy.Btu, UnitTypes.Energy);

            _conversions["bit"] = new Unit(InformationUnits.Bit, UnitTypes.InformationUnits);
            _conversions["byte"] = new Unit(InformationUnits.Byte, UnitTypes.InformationUnits);

            InitMetricPrefixes();
            InitBinaryPrefixes();

            _initialized = true;
        }

        #endregion

        #region Definitions
        public class Distance
        {
            public const double Meter = 1d;
            public const double StatueMile = 1609.344;
            public const double Nmi = 1852d;
            public const double Foot = 0.3048;
            public const double Inch = 0.0254;
            public const double Yard = 0.9144;
            public const double Angstrom = 0.0000000001;
            public const double Pica = 0.00423333333;
            public const double Picapt = 0.000352778;
            public const double Ell = 1.143;
            public static double LightYear = 9.46073 * System.Math.Pow(10, 15);
            public static double Parsec = 3.08567758 * System.Math.Pow(10, 16);
        }

        public class Time
        {
            public const double Hour = 1d;
            public const double Year = Day * 365.25;
            public const double Day = Hour * 24;
            public static double Minute = Hour/60d;
            public static double Second = Hour/3600d;
        }

        public class Speed
        {
            public const double AdmiralKnot = 0.51477333;
            public const double Knot = 0.514444444;
            public static double MetersPerHour = MetersPerSecond / 3600d;
            public const double MetersPerSecond = 1d;
            public const double MilesPerHour = 0.440704;
        }

        public class WeightAndMass
        {
            public const double Gram = 1d;
            public const double Slug = 14593.90294;
            public const double PoundMass = 453.59237;
            public static double U = 1.66054 * System.Math.Pow(10, -24);
            public const double OunceMass = 28.34952313;
            public const double Grain = 0.06479891;
            public const double UsHundredweight = 45359.237;
            public const double ImperialHundredweight = 50802.34544;
            public const double Stone = 6350.29318;
            public const double Ton = 907184.74;
            public const double ImperialTon = 1016046.909;
        }

        public class Area
        {
            public const double InternationalAcre = 4046.856422d;
            public const double UsServeyAcre = 4046.87261d;
            public static double SquareAngstrom = MathObj.Pow(10, -20);
            public const double Are = 100d;
            public const double SquareFeet = 0.09290304d;
            public const double Hectare = 10000d;
            public const double SquareInches = 0.00064516d;
            public static double SquareLightYear = MathObj.Pow(10, 31) * 8.95054;
            public const double SquareMeter = 1d;
            public const double Morgen = 2500d;
            public const double SquareMiles = 2589988.11d;
            public const double SquareNauticalMiles = 3429904d;
            public static double SquarePica = MathObj.Pow(10, -7) * 1.24452;
            public const double SquareYards = 0.83612736d;
        }

        public class Liquid
        {
            public const double Teaspoon = 0.00492892159d;
            public const double Tablespoon = 0.0147867648d;
            public const double FluidOunce = 0.0295735296d;
            public const double Cup = 0.236588237d;
            public const double UsPint = 0.473176473d;
            public const double UkPint = 0.56826125d;
            public const double Quart = 0.946352946d;
            public const double ImperialQuart = 1.1365225d;
            public const double Gallon = 3.785411784d;
            public const double Liter = 1d;
            public static double CubicAngstrom = MathObj.Pow(10, -27);
            public const double UsOilBarrel = 158.9872949d;
            public const double UsBushel = 35.23907017d;
            public const double CubicFeet = 28.31684659d;
            public const double CubicInch = 0.016387064d;
            public static readonly double CubicLightYear = MathObj.Pow(10, 50) * 8.46787;
            public const double CubicMeter = 1000d;
            public static readonly double CubicMile = MathObj.Pow(10, 12) * 4.16818;
            public const double CubicYard = 764.554858d;
            public static readonly double CubicNauticalMile = MathObj.Pow(10, 12) * 6.35218;
            public static readonly double CubicPica = MathObj.Pow(10, -8) * 4.3904;
            public const double GrossRegistredTon = 2831.684659;
            public const double MeasurementTon = 1132.673864;
        }

        public class Power
        {
            public const double Horsepower = 745.6998716d;
            public const double Watt = 1d;
            public const double PS = 735.49875d;
        }

        public class Magnetism
        {
            public const double Tesla = 10000d;
            public const double Gauss = 1d;
        }

        public class Force
        {
            public const double Newton = 1d;
            public const double Dyne = 0.00001d;
            public const double PoundForce = 4.448221615d;
            public const double Pond = 0.00980665d;
        }

        public class Pressure
        {
            public const double Pascal = 1d;
            public const double Atmosphere = 101325d;
            public const double MmOfMercury = 133.322d;
            public const double Psi = 6894.757293d;
            public const double Torr = 133.3223684d;
        }

        public class Temperature
        {
            public const double Celcius = 1d;
            public const double Fahrenheit = 1d;
            public const double Kelvin = 1d;
            public const double Rankine = 1d;
            public const double Reaumur = 1d;
        }

        public class Energy
        {
            public const double Joule = 1d;
            public const double Erg = 0.0000001d;
            public const double ThermodynamicCalorie = 4.184d;
            public const double ItCalorie = 4.1868d;
            public static double ElectronVolt = 1.60217662 * MathObj.Pow(10, -19);
            public static double HorsePowerHour = 2684519.538;
            public static double WattHour = 3600;
            public static double FootPound = 1.3558179483314;
            public static double Btu = 1055.0558526;
        }

        public class InformationUnits
        {
            public const double Bit = 1d;
            public const double Byte = 8d;
        }

        #endregion

        #region Public methods

        public static bool IsValidUnit(string unit)
        {
            lock (_syncRoot)
            {
                if (!_initialized) Init();
            }
            if (!_conversions.ContainsKey(unit) && !TemperatureConverter.IsValidUnit(unit))
            {
                return false;
            }
            return true;
        }

        public static double Convert(double number, string fromUnit, string toUnit)
        {
            lock (_syncRoot)
            {
                if (!_initialized) Init();
            }
            // special case for temperatures which has to be converted with a function in some cases
            if (TemperatureConverter.IsTempMapping(fromUnit, toUnit))
            {
                return TemperatureConverter.Conversions[$"{fromUnit}>{toUnit}"](number);
            }

            if (!IsValidUnit(fromUnit))
            {
                throw new ArgumentException("Invalid fromUnit", fromUnit);
            }
            if(!IsValidUnit(toUnit))
            {
                throw new ArgumentException("Invalid toUnit", toUnit);
            }
            var from = _conversions[fromUnit];
            var to = _conversions[toUnit];
            if (from.UnitType != to.UnitType)
            {
                throw new ArgumentException("Units are not compatible: " + fromUnit + " and " + toUnit);
            }
            return number * (from.Value / to.Value);
        }

        #endregion
    }
}
