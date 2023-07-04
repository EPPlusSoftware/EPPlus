using System;
using System.Collections.Generic;
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
            if(string.IsNullOrEmpty(_input) || _input.Length < 5)
            {
                IsValidFormat= false;
                SerialNumber = double.NaN;
                return;
            }
            var hour = 0;
            var minute = 0;
            var second = 0;
            var isValid = false;
            if(_input.Contains(":"))
            {
                var array = _input.Split(':');
                if (int.TryParse(array[0], out int h) && int.TryParse(array[1], out int m)) 
                {
                    hour = h;
                    minute = m;
                    isValid= true;
                }
                if(isValid && array.Length > 2)
                {
                    if (int.TryParse(array[2], out int s))
                    {
                        second = s;
                    }
                    else
                    {
                        isValid = false;
                    }
                }
                if(isValid)
                {
                    var arr = _input.Split(' ');
                    if(arr.Length > 2)
                    {
                        isValid = false;
                    }
                    else if(arr.Length == 2)
                    {
                        var ampmPart = arr[1].Trim();
                        if(string.Compare(ampmPart, "AM", true) == 0 || string.Compare(ampmPart, "PM", true) == 0)
                        {
                            AmPm = ampmPart;
                        }
                        else
                        {
                            isValid = false;
                        }
                    }
                }
                if(isValid)
                {
                    isValid = AreValidTimeValues(minute, second);
                    if(isValid)
                    {
                        SerialNumber = GetSerialNumber(hour, minute, second);
                    }
                }
                else
                {
                    if(DateTime.TryParse(_input, out DateTime dt))
                    {
                        SerialNumber = GetSerialNumber(dt.Hour, dt.Minute, dt.Second);
                        isValid = true;
                    }
                    else
                    {
                        isValid = false;
                    }
                }
                if(!isValid)
                {
                    SerialNumber = double.NaN;
                }
                IsValidFormat = isValid;
            }
        }

        private double GetSerialNumber(int hour, int minute, int second)
        {
            var secondsInADay = 24d * 60d * 60d;
            return ((double)hour * 60 * 60 + (double)minute * 60 + (double)second) / secondsInADay;
        }

        private bool AreValidTimeValues(int minute, int second)
        {
            if (second < 0 || second > 59)
            {
                return false;
            }
            if (minute < 0 || minute > 59)
            {
                return false;
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
