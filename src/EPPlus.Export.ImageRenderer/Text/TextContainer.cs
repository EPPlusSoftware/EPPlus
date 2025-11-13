using EPPlus.Fonts.OpenType;
using EPPlusImageRenderer.Text;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.ImageRenderer.Text
{
    internal class TextContainer : RectMargins
    {
        /// <summary>
        /// Should we resize the container depending on the text
        /// </summary>
        bool ResizeToFit = false;

        List<string> _textContentCollection = new List<string>();

        MeasurementFont MeasurementFont;
        FontMeasurerTrueType _textMeasurerTrueType = null;

        public bool AddDeltaHeight = false;

        List<string> TextContent
        {
            get
            {
                return _textContentCollection;
            }
            set
            {
                _textContentCollection = value;
                if (ResizeToFit)
                {
                    AutofitContainer();
                }
            }
        }


        public RectBase GetTextArea()
        {
            return GetInnerRect();
        }


        //To create a placeholder with default values
        public TextContainer()  : base() 
        {
            ResizeToFit = false;
            MeasurementFont = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Style = MeasurementFontStyles.Regular,
                Size = 11
            };

            //Excel at 100% size appears to have this pixel right and bottom for cells
            Left = 0; Top = 0; Right = 64; Bottom = 20d;

            //Initalize margins to 0
            MarginLeft = MarginRight = MarginTop = MarginBottom = 0;

            _textContentCollection = new List<string>();
        }

        /// <summary>
        /// </summary>
        /// <param name="text"></param>
        /// <param name="font"></param>
        /// <param name="resizeToFit"></param>
        public TextContainer(List<string> textCollection, MeasurementFont font, bool resizeToFit = false) 
        { 
            MeasurementFont = font;
            ResizeToFit = resizeToFit;
            TextContent = textCollection;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="textCollection"></param>
        /// <param name="font"></param>
        /// <param name="resizeToFit"></param>
        public TextContainer(string text, MeasurementFont font, bool resizeToFit = false, bool addDeltaHeight = false)
        {
            MeasurementFont = font;
            AddDeltaHeight = addDeltaHeight;
            ResizeToFit = resizeToFit;
            TextContent = SplitStrings(text);
        }

        public static readonly string[] lineEndings = { "\r\n", "\r", "\n" };

        List<string> SplitStrings(string text)
        {
            string[] splitLines = text.Split
            (
                lineEndings, 
                StringSplitOptions.None
            );
            return splitLines.ToList();
        }

        void AutofitContainer()
        {
            if (_textMeasurerTrueType == null)
            {
                _textMeasurerTrueType = new FontMeasurerTrueType(MeasurementFont);
            }

            double maxWidth = 0;
            foreach ( var line in TextContent)
            {
                var lineWidth = _textMeasurerTrueType.MeasureTextWidthInPixels(line);
                var lineWidthAlt = _textMeasurerTrueType.MeasureTextWidthInPixels(line+"\r\n");
                //lineWidth += _textMeasurerTrueType.GetMinXInPixels();
                if (lineWidth > maxWidth)
                {
                    maxWidth = lineWidth;
                }
            }
            var totalHeight = _textMeasurerTrueType.GetHeightOfTextInPixels(_textContentCollection);

            //if(AddDeltaHeight)
            //{
            //    //Text boxes in excel have a slight additional pixel size roughly equal to difference between typoAscent and winAscent
            //    var deltaAscent = _textMeasurerTrueType.GetDeltaAscent(TextUnit.Pixels);
            //    totalHeight += deltaAscent;
            //}

            Width = maxWidth;
            Height = totalHeight;
        }
    }
}
