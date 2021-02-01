using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils
{
    internal static class CsvExtensions
    {
        public static string GetCsvPosition(this string argument, int position) 
        {
            if (string.IsNullOrEmpty(argument)) return "";
            var items = argument.Split(',');
            if(items.Length > position)
            {
                return items[position];
            }
            return "";
        }
        public static string SetCsvPosition(this string argument, int position,int size, string value, string defaultValue="0")
        {
            if (argument == null) return value;
            var items = argument.Split(',');
            if(items.Length < size)
            {
                var newItems = new string[size];
                Array.Copy(items, newItems, items.Length);
                for(int i=items.Length;i<size;i++)
                {
                    newItems[i] = defaultValue;
                }
                items = newItems;
            }
            if (items.Length > position)
            {
                items[position]=value;
            }
            else
            {
                throw(new InvalidOperationException("CSV Position out our range"));
            }
            return string.Join(",", items);
        }
    }
}
