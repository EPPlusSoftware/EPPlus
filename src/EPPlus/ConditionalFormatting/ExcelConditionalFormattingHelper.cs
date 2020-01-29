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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.Utils;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
  /// <summary>
  /// Conditional formatting helper
  /// </summary>
  internal static class ExcelConditionalFormattingHelper
  {
    /// <summary>
    /// Check and fix an address (string address)
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public static string CheckAndFixRangeAddress(
      string address)
    {
      if (address.Contains(','))
      {
        throw new FormatException(
          ExcelConditionalFormattingConstants.Errors.CommaSeparatedAddresses);
      }

      address = ConvertUtil._invariantTextInfo.ToUpper(address);

      if (Regex.IsMatch(address, @"[A-Z]+:[A-Z]+"))
      {
        address = AddressUtility.ParseEntireColumnSelections(address);
      }

      return address;
    }
    /// <summary>
    /// Convert a color code to Color Object
    /// </summary>
    /// <param name="colorCode">Color Code (Ex. "#FFB43C53" or "FFB43C53")</param>
    /// <returns></returns>
    public static Color ConvertFromColorCode(
      string colorCode)
    {
      try
      {
        return Color.FromArgb(Int32.Parse(colorCode.Replace("#", ""), NumberStyles.HexNumber));
      }
      catch
      {
        // Assume white is the default color (instead of giving an error)
        return Color.White;
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="node"></param>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static string GetAttributeString(
      XmlNode node,
      string attribute)
    {
      try
      {
        var value = node.Attributes[attribute].Value;
        return (value == null) ? string.Empty : value;
      }
      catch
      {
        return string.Empty;
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="node"></param>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static int GetAttributeInt(
      XmlNode node,
      string attribute)
    {
      try
      {
        var value = node.Attributes[attribute].Value;
        return int.Parse(value, NumberStyles.Integer, CultureInfo.InvariantCulture);
      }
      catch
      {
        return int.MinValue;
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="node"></param>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static int? GetAttributeIntNullable(
      XmlNode node,
      string attribute)
    {
      try
      {
          if (node.Attributes[attribute] == null)
          {
              return null;
          }
          else
          {
              var value = node.Attributes[attribute].Value;
              return int.Parse(value, NumberStyles.Integer, CultureInfo.InvariantCulture);
          }
      }
      catch
      {
        return null;
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="node"></param>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static bool GetAttributeBool(
      XmlNode node,
      string attribute)
    {
      try
      {
        var value = node.Attributes[attribute].Value;
        return (value == "1" || value == "-1" || value.Equals("TRUE", StringComparison.OrdinalIgnoreCase));
      }
      catch
      {
        return false;
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="node"></param>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static bool? GetAttributeBoolNullable(
      XmlNode node,
      string attribute)
    {
      try
      {
          if (node.Attributes[attribute] == null)
          {
              return null;
          }
          else
          {
              var value = node.Attributes[attribute].Value;
              return (value == "1" || value == "-1" || value.Equals("TRUE",StringComparison.OrdinalIgnoreCase));
          }
      }
      catch
      {
        return null;
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="node"></param>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static double GetAttributeDouble(
      XmlNode node,
      string attribute)
    {
      try
      {
        var value = node.Attributes[attribute].Value;
        return double.Parse(value, NumberStyles.Number, CultureInfo.InvariantCulture);
      }
      catch
      {
        return double.NaN;
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="node"></param>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static decimal GetAttributeDecimal(
      XmlNode node,
      string attribute)
    {
      try
      {
        var value = node.Attributes[attribute].Value;
        return decimal.Parse(value, NumberStyles.Any, CultureInfo.InvariantCulture);
      }
      catch
      {
        return decimal.MinValue;
      }
    }

    /// <summary>
    /// Encode to XML (special characteres: &apos; &quot; &gt; &lt; &amp;)
    /// </summary>
    /// <param name="s"></param>
    /// <returns></returns>
    public static string EncodeXML(
      this string s)
    {
      return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
    }

    /// <summary>
    /// Decode from XML (special characteres: &apos; &quot; &gt; &lt; &amp;)
    /// </summary>
    /// <param name="s"></param>
    /// <returns></returns>
    public static string DecodeXML(
      this string s)
    {
      return s.Replace("'", "&apos;").Replace("\"", "&quot;").Replace(">", "&gt;").Replace("<", "&lt;").Replace("&", "&amp;");
    }
  }
}