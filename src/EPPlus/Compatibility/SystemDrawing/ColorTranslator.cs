//
// System.Drawing.ColorTranslator.cs
//
// Authors:
//	Dennis Hayes (dennish@raytek.com)
//	Ravindra (rkumar@novell.com)
//	Sebastien Pouliot  <sebastien@ximian.com>
//
// Copyright (C) 2001 Ximian, Inc.  http://www.ximian.com
// Copyright (C) 2004,2006-2007 Novell, Inc (http://www.novell.com)
//
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
// 
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//

using OfficeOpenXml.Drawing.Style.Coloring;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;

namespace OfficeOpenXml.Compatibility.System.Drawing
{

	internal sealed class ColorTranslator
	{

		private ColorTranslator()
		{
		}

		public static Color FromHtml(string htmlColor)
		{
			if (string.IsNullOrEmpty(htmlColor)) return Color.Empty;

			switch (htmlColor.ToLowerInvariant())
			{
				case "buttonface":
				case "threedface":
					return EPPlusSystemColors.Control;
				case "buttonhighlight":
				case "threedlightshadow":
					return EPPlusSystemColors.ControlLightLight;
				case "buttonshadow":
					return EPPlusSystemColors.ControlDark;
				case "captiontext":
					return EPPlusSystemColors.ActiveCaptionText;
				case "threeddarkshadow":
					return EPPlusSystemColors.ControlDarkDark;
				case "threedhighlight":
					return EPPlusSystemColors.ControlLight;
				case "background":
					return EPPlusSystemColors.Desktop;
				case "buttontext":
					return EPPlusSystemColors.ControlText;
				case "infobackground":
					return EPPlusSystemColors.Info;
				// special case for Color.LightGray versus html's LightGrey (#340917)
				case "lightgrey":
					return Color.LightGray;
			}

			if (htmlColor[0] == '#' && htmlColor.Length == 4)
			{
				char r = htmlColor[1], g = htmlColor[2], b = htmlColor[3];
				htmlColor = new string(new char[] { '#', r, r, g, g, b, b });
			}

			TypeConverter converter = TypeDescriptor.GetConverter(typeof(Color));
			return (Color)converter.ConvertFromString(htmlColor);
		}
	}
}