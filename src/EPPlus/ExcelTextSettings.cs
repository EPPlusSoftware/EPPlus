﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.SystemDrawing.Text;
using System;

namespace OfficeOpenXml
{
    /// <summary>
    /// This class contains settings for text measurement.
    /// </summary>
    public class ExcelTextSettings
    {
        internal ExcelTextSettings()
        {
            if(Environment.OSVersion.Platform==PlatformID.Unix ||
               Environment.OSVersion.Platform==PlatformID.MacOSX)
            {
                PrimaryTextMeasurer = new GenericFontMetricsTextMeasurer();
                try
                {
                    FallbackTextMeasurer = new SystemDrawingTextMeasurer();
                }
                catch
                {
                    FallbackTextMeasurer = null;
                }
            }
            else
            {
                try
                {
                    var m = new SystemDrawingTextMeasurer();
                    if (m.ValidForEnvironment())
                    {
                        PrimaryTextMeasurer = m;
                        FallbackTextMeasurer = new GenericFontMetricsTextMeasurer();
                    }
                    else
                    {
                        PrimaryTextMeasurer = new GenericFontMetricsTextMeasurer();
                    }
                }
                catch
                {
                    PrimaryTextMeasurer = new GenericFontMetricsTextMeasurer();
                }
            }
            AutofitScaleFactor = 1f;
        }

        /// <summary>
        /// This is the primary text measurer
        /// </summary>
        public ITextMeasurer PrimaryTextMeasurer { get; set; }

        /// <summary>
        /// If the primary text measurer fails to measure the text, this one will be used.
        /// </summary>
        public ITextMeasurer FallbackTextMeasurer { get; set; }

        /// <summary>
        /// All measurements of texts will be multiplied with this value. Default is 1.
        /// </summary>
        public float AutofitScaleFactor { get; set; }
        /// <summary>
        /// A percentage of the widest text. Since charaters in different fonts have different widths we use this threshold remove characters from the longer string for comparing to the current text.
        /// This is so we can skip obvious shorter strings and save time on calculating it's actual width.
        /// </summary>
        public double textLengthThreshold = 0.5d;
        /// <summary>
        /// The ammount of rows to check for when autofitting, starts from top. A value set to 0 or lower means checking all rows in the column.
        /// </summary>
        public int AutofitRows = 0;

        /// <summary>
        /// Returns an instance of the internal generic text measurer
        /// </summary>
        public ITextMeasurer GenericTextMeasurer
        {
            get
            {
                return new GenericFontMetricsTextMeasurer();
            }
        }

        /// <summary>
        /// Measures a text with default settings when there is no other option left...
        /// </summary>
        internal DefaultTextMeasurer DefaultTextMeasurer { get; set; }
    }
}
