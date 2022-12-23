/*************************************************************************************************
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
            _autofitWidthScaleFactor = 1f;
            _autofitHeightScaleFactor = 1f;
        }

        /// <summary>
        /// This is the primary text measurer
        /// </summary>
        public ITextMeasurer PrimaryTextMeasurer { get; set; }

        /// <summary>
        /// If the primary text measurer fails to measure the text, this one will be used.
        /// </summary>
        public ITextMeasurer FallbackTextMeasurer { get; set; }

        private float _autofitWidthScaleFactor;

        /// <summary>
        /// All measurements of text-width will be multiplied with this value. Default is 1.
        /// </summary>
        [Obsolete("Will be removed in future major versions. Use AutofitWidthScaleFactor or AutofitHeightScaleFactor instead.")]
        public float AutofitScaleFactor
        {
            get
            {
                return _autofitWidthScaleFactor;
            }
            set
            {
                if (value < 0.3f || value > 3f) throw new ArgumentException("AutofitScaleFactor: value must be between 0.3 and 3");
                _autofitWidthScaleFactor = value;
            }
        }

        /// <summary>
        /// All measurements of text-width will be multiplied with this value. Default is 1.
        /// </summary>
        public float AutofitWidthScaleFactor
        {
            get
            {
                return _autofitWidthScaleFactor;
            }
            set
            {
                if (value < 0.3f || value > 3f) throw new ArgumentException("AutofitWidthScaleFactor: value must be between 0.3 and 3");
                _autofitWidthScaleFactor = value;
            }
        }

        private float _autofitHeightScaleFactor;

        /// <summary>
        /// All measurements of text-height will be multiplied with this value. Default is 1.
        /// </summary>
        public float AutofitHeightScaleFactor
        {
            get
            {
                return _autofitHeightScaleFactor;
            }
            set
            {
                if (value < 0.3f || value > 3f) throw new ArgumentException("AutofitHeightScaleFactor: value must be between 0.3 and 3");
                _autofitHeightScaleFactor = value;
            }
        }
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
