using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime
{
    internal class TimeString
    {
        public TimeString(string input)
        {
            _input = input;
            Initialize();
        }

        private readonly string _input;

        private void Initialize()
        {
            if(string.IsNullOrEmpty(_input) || _input.Length < 3)
            {
                IsValidFormat= false;
                SerialNumber = double.NaN;
                return;
            }
            var input = SetAmPmPart(_input);
            var hour = 0;
            var minute = 0;
            var second = 0d;
            var isValid = false;
            if(!input.Contains(":"))
            {
                var arr = _input.Split(' ');
                if (arr.Length == 2 && int.TryParse(arr[0], out int h) && !string.IsNullOrEmpty(arr[1]))
                {
                    if((AmPm == "AM" || AmPm == "PM") && h >=0 || h <= 23)
                    {
                        isValid = true;
                        hour = h;
                        IsValidFormat = true;
                    }
                }
            }
            else
            {
                // input didn't contain ':'
                var array = _input.Split(':');
                if (array.Length >=1)
                {
                    isValid = int.TryParse(array[0], out int h);
                    if(isValid)
                    {
                        hour = h;
                    }
                }
                if (isValid && array.Length > 1) 
                {
                    isValid = int.TryParse(array[1], out int m);
                    if(isValid)
                    {
                        minute = m;
                    }
                }

                if(isValid && array.Length > 2)
                {
                    isValid = IsValidDouble(array[2]);
                    if (isValid)
                    {
                        second = double.Parse(array[2], CultureInfo.InvariantCulture);
                    }
                }
            }
            if (isValid)
            {
                isValid = AreValidTimeValues(hour, minute, second);
                if (isValid)
                {
                    SerialNumber = GetSerialNumber(hour, minute, second);
                }
            }
            else
            {
                if (DateTime.TryParse(_input, out DateTime dt))
                {
                    SerialNumber = GetSerialNumber(dt.Hour, dt.Minute, dt.Second);
                    isValid = true;
                }
                else
                {
                    isValid = false;
                }
            }
            if (!isValid)
            {
                SerialNumber = double.NaN;
            }
            IsValidFormat = isValid;
        }

        private string SetAmPmPart(string input)
        {
            var inp = input.Trim().ToUpperInvariant();
            if(inp.EndsWith("AM"))
            {
                inp = inp.Substring(0, inp.Length - 2);
                AmPm = "AM";
            }
            else if(inp.EndsWith("PM"))
            {
                inp = inp.Substring(0, inp.Length -2);
                AmPm = "PM";
            }
            return inp;
        }

        private double GetSerialNumber(int hour, int minute, double second)
        {
            var secondsInADay = 24d * 60d * 60d;
            return ((double)hour * 60 * 60 + (double)minute * 60 + second) / secondsInADay;
        }

        private bool AreValidTimeValues(int hour, int minute, double second)
        {
            if (second < 0 || second > 59)
            {
                return false;
            }
            if (minute < 0 || minute > 59)
            {
                return false;
            }
            if(!string.IsNullOrEmpty(AmPm))
            {
                if(hour > 11 && AmPm == "AM")
                {
                    return false;
                }
            }
            return true;
        }

        private static bool IsValidDouble(string d)
        {
            if(string.IsNullOrEmpty(d))
            {
                return false;
            }
            foreach(var c in d)
            {
                if(!char.IsDigit(c) && c != '.')
                {
                    return false;
                }
            }
            return true;
        }

        public int Hour { get; private set; }

        public int Minute { get; private set;}

        public int Second { get; private set; }

        public string AmPm { get; private set; }

        public double SerialNumber { get; private set; }

        public bool IsValidFormat { get; private set; }
    }
}
