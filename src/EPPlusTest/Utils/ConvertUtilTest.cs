/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using System;
using System.ComponentModel;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Compatibility;

namespace EPPlusTest.Utils
{
	[TestClass]
	public class ConvertUtilTest
	{
		[TestMethod]
		public void TryParseNumericString()
		{
			double result;
			object numericString = null;
			double expected = 0;
			Assert.IsFalse(ConvertUtil.TryParseNumericString(numericString as string, out result));
			Assert.AreEqual(expected, result);
			expected = 1442.0;
			numericString = expected.ToString("e", CultureInfo.CurrentCulture); // 1.442E+003
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString as string, out result));
			Assert.AreEqual(expected, result);
			numericString = expected.ToString("f0", CultureInfo.CurrentCulture); // 1442
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString as string, out result));
			Assert.AreEqual(expected, result);
			numericString = expected.ToString("f2", CultureInfo.CurrentCulture); // 1442.00
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString as string, out result));
			Assert.AreEqual(expected, result);
			numericString = expected.ToString("n", CultureInfo.CurrentCulture); // 1,442.0
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString as string, out result));
			Assert.AreEqual(expected, result);
			expected = -0.00526;
			numericString = expected.ToString("e", CultureInfo.CurrentCulture); // -5.26E-003
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString as string, out result));
			Assert.AreEqual(expected, result);
			numericString = expected.ToString("f0", CultureInfo.CurrentCulture); // -0
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString as string, out result));
			Assert.AreEqual(0.0, result);
			numericString = expected.ToString("f3", CultureInfo.CurrentCulture); // -0.005
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString as string, out result));
			Assert.AreEqual(-0.005, result);
			numericString = expected.ToString("n6", CultureInfo.CurrentCulture); // -0.005260
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString as string, out result));
			Assert.AreEqual(expected, result);
		}
		
		[TestMethod]
		public void TryParseDateString()
		{
			DateTime result;
			object dateString = null;
			DateTime expected = DateTime.MinValue;
			Assert.IsFalse(ConvertUtil.TryParseDateString(dateString as string, out result));
			Assert.AreEqual(expected, result);
			expected = new DateTime(2013, 1, 15);
			dateString = expected.ToString("d", CultureInfo.CurrentCulture); // 1/15/2013
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString as string, out result));
			Assert.AreEqual(expected, result);
			dateString = expected.ToString("D", CultureInfo.CurrentCulture); // Tuesday, January 15, 2013
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString as string, out result));
			Assert.AreEqual(expected, result);
			dateString = expected.ToString("F", CultureInfo.CurrentCulture); // Tuesday, January 15, 2013 12:00:00 AM
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString as string, out result));
			Assert.AreEqual(expected, result);
			dateString = expected.ToString("g", CultureInfo.CurrentCulture); // 1/15/2013 12:00 AM
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString as string, out result));
			Assert.AreEqual(expected, result);
			expected = new DateTime(2013, 1, 15, 15, 26, 32);
			dateString = expected.ToString("F", CultureInfo.CurrentCulture); // Tuesday, January 15, 2013 3:26:32 PM
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString as string, out result));
			Assert.AreEqual(expected, result);
			dateString = expected.ToString("g", CultureInfo.CurrentCulture); // 1/15/2013 3:26 PM
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString as string, out result));
			Assert.AreEqual(new DateTime(2013, 1, 15, 15, 26, 0), result);
		}

        
        // This is just illustration of the bug in old implementation
        //[TestMethod]
        public void TextToIntInOldImplementation()
        {
            var result = GetTypedValue<int>("204");

            Assert.AreEqual(204, result);
        }

        // previous implementation
        internal T GetTypedValue<T>(object v)
        {
            if (v == null)
            {
                return default(T);
            }
            Type fromType = v.GetType();
            Type toType = typeof(T);
            
            Type toType2 = (TypeCompat.IsGenericType(toType) && toType.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
                ? Nullable.GetUnderlyingType(toType)
                : null;
            if (fromType == toType || fromType == toType2)
            {
                return (T)v;
            }
            var cnv = TypeDescriptor.GetConverter(fromType);
            if (toType == typeof(DateTime) || toType2 == typeof(DateTime))    //Handle dates
            {
                if (fromType == typeof(TimeSpan))
                {
                    return ((T)(object)(new DateTime(((TimeSpan)v).Ticks)));
                }
                else if (fromType == typeof(string))
                {
                    DateTime dt;
                    if (DateTime.TryParse(v.ToString(), out dt))
                    {
                        return (T)(object)(dt);
                    }
                    else
                    {
                        return default(T);
                    }

                }
                else
                {
                    if (cnv.CanConvertTo(typeof(double)))
                    {
                        return (T)(object)(DateTime.FromOADate((double)cnv.ConvertTo(v, typeof(double))));
                    }
                    else
                    {
                        return default(T);
                    }
                }
            }
            else if (toType == typeof(TimeSpan) || toType2 == typeof(TimeSpan))    //Handle timespan
            {
                if (fromType == typeof(DateTime))
                {
                    return ((T)(object)(new TimeSpan(((DateTime)v).Ticks)));
                }
                else if (fromType == typeof(string))
                {
                    TimeSpan ts;
                    if (TimeSpan.TryParse(v.ToString(), out ts))
                    {
                        return (T)(object)(ts);
                    }
                    else
                    {
                        return default(T);
                    }
                }
                else
                {
                    if (cnv.CanConvertTo(typeof(double)))
                    {

                        return (T)(object)(new TimeSpan(DateTime.FromOADate((double)cnv.ConvertTo(v, typeof(double))).Ticks));
                    }
                    else
                    {
                        try
                        {
                            // Issue 14682 -- "GetValue<decimal>() won't convert strings"
                            // As suggested, after all special cases, all .NET to do it's 
                            // preferred conversion rather than simply returning the default
                            return (T)Convert.ChangeType(v, typeof(T));
                        }
                        catch (Exception)
                        {
                            // This was the previous behaviour -- no conversion is available.
                            return default(T);
                        }
                    }
                }
            }
            else
            {
                if (cnv.CanConvertTo(toType))
                {
                    return (T)cnv.ConvertTo(v, typeof(T));
                }
                else
                {
                    if (toType2 != null)
                    {
                        toType = toType2;
                        if (cnv.CanConvertTo(toType))
                        {
                            return (T)cnv.ConvertTo(v, toType); //Fixes issue 15377
                        }
                    }

                    if (fromType == typeof(double) && toType == typeof(decimal))
                    {
                        return (T)(object)Convert.ToDecimal(v);
                    }
                    else if (fromType == typeof(decimal) && toType == typeof(double))
                    {
                        return (T)(object)Convert.ToDouble(v);
                    }
                    else
                    {
                        return default(T);
                    }
                }
            }
        }

    }
}
