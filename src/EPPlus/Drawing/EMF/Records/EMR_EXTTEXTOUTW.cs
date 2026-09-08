/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2025         EPPlus Software AB           Initial release EPPlus 8
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.GenericFontWidths;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Utils.FileUtils;
using System;
using System.IO;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_EXTTEXTOUTW : EMR_RECORD
    {
        internal byte[] Bounds;
        internal byte[] iGraphicsMode;
        internal byte[] exScale;
        internal byte[] eyScale;
        internal PointL Reference;
        internal uint Chars;
        internal uint offString;
        internal byte[] Options;
        internal byte[] Rectangle;
        internal uint offDx;
        internal string stringBuffer;
        internal byte[] DxBuffer;

        private int padding = 0;

        internal uint InternalFontId;
        internal ExcelTextSettings textSettings = new ExcelTextSettings();

        internal MapMode mode = MapMode.MM_TEXT;
        internal ITextMeasurer Measurer;

        internal float Ppi = 108.73578912433f;
        internal float UnitsPerEm = 2295f;

        //string pixelWidth
        internal uint? totalPixelWidthChars = null;

        /// <summary>
        /// Minimum spacing is 0x01 which should be correct at fontsize 2
        /// </summary>
        //internal int FontSize = 11;
        internal int FontPointSize
        {
            get
            {
                if (Font == null | Font.elw.Height == 0)
                {
                    return 11;
                }
                else
                {
                    var height = Font.elw.Height;

                    return Font.elw.Height < 0 ? Math.Abs(height) : height;
                }
            }
        }

        internal EMR_EXTCREATEFONTINDIRECTW Font = null;

        internal string Text
        {
            get
            {
                return stringBuffer;
            }
            set
            {
                //var test = FontSize.GetFontSize(Font.elw.FaceName, true);
                //textSettings.GenericTextMeasurer.MeasureText(value, Meas)
                stringBuffer = value;
                CalculateOffsets();
            }
        }

        internal EMR_EXTTEXTOUTW(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            Bounds = br.ReadBytes(16);
            iGraphicsMode = br.ReadBytes(4);
            exScale = br.ReadBytes(4);
            eyScale = br.ReadBytes(4);
            Reference = new PointL(br);        //Signed, koordinat för var texten börjar. 
            Chars = br.ReadUInt32();
            offString = br.ReadUInt32();
            Options = br.ReadBytes(4);
            Rectangle = br.ReadBytes(16);
            offDx = br.ReadUInt32();
            br.BaseStream.Position = position + offString;
            //Encoding is specifically UTF-16LE meaning no BOM allowed and little endian
            stringBuffer = BinaryHelper.GetString(br, (Chars * 2), Encoding.Unicode);
            br.BaseStream.Position = position + offDx;
            DxBuffer = br.ReadBytes((int)(Size - offDx));

            var changedSize = offDx - offString;
            changedSize -= (Chars * 2);
            if (changedSize > 0)
            {
                padding = (int)changedSize;
            }
        }

        internal byte[] CalculateDxSpacing(string targetString)
        {
            var aMesurement = (GenericFontMetricsTextMeasurer)textSettings.GenericTextMeasurer;
            aMesurement.MeasureTextInternal(targetString, GenericTextMeasurerKey.GetKey(Font.elw.mFont.FontFamily, Font.elw.mFont.Style), Font.elw.mFont.Style, Font.elw.mFont.Size);
            var values = aMesurement.MeasureIndividualCharacters(targetString, Font.elw.mFont, Ppi);

            var measurement = aMesurement.MeasureText(targetString, Font.elw.mFont);

            int index = 0;
            uint sum = 0;

            foreach (uint val in values)
            {
                var bytes = BitConverter.GetBytes(val);
                bytes.CopyTo(DxBuffer, index);
                index += bytes.Length;
                sum += val;
            }

            totalPixelWidthChars = sum;
            return DxBuffer;
        }

        /// <summary>
        /// Calculate centering for text
        /// </summary>
        /// <param name="totalWidth">Total width (x value endpoint in emf world)</param>
        /// <param name="minWidth">Total width (x value endpoint in emf world)</param>
        internal void AdjustReferenceToCenterText(int totalWidth, int minWidth)
        {
            if(totalPixelWidthChars == null)
            {
                CalculateDxSpacing(stringBuffer);
            }

            var point = Convert.ToInt32((totalWidth - totalPixelWidthChars) * 0.5);

            point = point < minWidth ? minWidth : point;

            Reference.PointX = point;
        }

        internal int GetSpacingForChar(char c)
        {
            return GetSpacingForChar(c, (GenericFontMetricsTextMeasurer)textSettings.GenericTextMeasurer, Font.elw.mFont, Ppi);
        }

        internal static int GetSpacingForChar(char c, GenericFontMetricsTextMeasurer aMesurement, MeasurementFont mFont, float ppi)
        {
            return (int)aMesurement.MeasureIndividualCharacter(c, mFont, ppi);
        }

        internal EMR_EXTTEXTOUTW(string Text, EMR_EXTCREATEFONTINDIRECTW font)
        {
            Font = font;

            Type = RECORD_TYPES.EMR_EXTTEXTOUTW;
            Bounds = new byte[16] { 0x13, 0x00, 0x00, 0x00, 0x24, 0x00, 0x00, 0x00, 0x4b, 0x00, 0x00, 0x00, 0x30, 0x00, 0x00, 0x00 };
            iGraphicsMode = new byte[4] { 0x02, 0x00, 0x00, 0x00 };
            exScale = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            eyScale = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            //Reference = new byte[8] { 0x13, 0x00, 0x00, 0x00, 0x24, 0x00, 0x00, 0x00 };
            Reference = new PointL(19,36);
            Options = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            Rectangle = new byte[16] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF };
            offString = 4 + 4 + 16 + 4 + 4 + 4 + 8 + 4 + 4 + 4 + 16 + 4;
            stringBuffer = Text;
            CalculateOffsets();
        }

        internal EMR_EXTTEXTOUTW(string Text, int x, int y, EMR_EXTCREATEFONTINDIRECTW font)
        {
            Font = font;

            Type = RECORD_TYPES.EMR_EXTTEXTOUTW;
            Bounds = new byte[16] { 0x13, 0x00, 0x00, 0x00, 0x24, 0x00, 0x00, 0x00, 0x4b, 0x00, 0x00, 0x00, 0x30, 0x00, 0x00, 0x00 };
            iGraphicsMode = new byte[4] { 0x02, 0x00, 0x00, 0x00 };
            exScale = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            eyScale = new byte[4] { 0x00, 0x00, 0x00, 0x00 };

            Reference = new PointL(x, y);

            //Reference = new byte[8] { 0x13, 0x00, 0x00, 0x00, 0x24, 0x00, 0x00, 0x00 };
            Options = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            Rectangle = new byte[16] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF };
            offString = 4 + 4 + 16 + 4 + 4 + 4 + 8 + 4 + 4 + 4 + 16 + 4;
            stringBuffer = Text;

            CalculateOffsets();
        }

        private void CalculateOffsets()
        {
            if(stringBuffer == null)
            {
                stringBuffer = "";
            }

            Chars = (uint)stringBuffer.Length;
            offDx = offString + (uint)stringBuffer.Length * 2;

            padding = (int)offDx;
            offDx += 4 - (offDx % 4);
            padding = (int)(offDx) - padding;

            DxBuffer = new byte[stringBuffer.Length * 4];
            CalculateDxSpacing(stringBuffer);
            Size = offDx + (uint)DxBuffer.Length;
        }

        //private int RightRectangleX()
        //{
        //    //var rightBytes = new byte[] { Rectangle[8], Rectangle[9], Rectangle[10], Rectangle[11] };
        //    //BitConverter.ToInt32(rightBytes, 8);
        //    int testStuff = BitConverter.ToInt32(Bounds, 8);
        //    return testStuff;
        //}

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(Bounds);
            bw.Write(iGraphicsMode);
            bw.Write(exScale);
            bw.Write(eyScale);
            Reference.WriteBytes(bw);
            bw.Write(Chars);
            bw.Write(offString);
            bw.Write(Options);
            bw.Write(Rectangle);
            bw.Write(offDx);
            bw.Write(BinaryHelper.GetByteArray(stringBuffer, Encoding.Unicode));
            if (padding > 0)
            {
                bw.Write(new byte[padding]);
            }
            bw.Write(DxBuffer);
        }
    }
}
