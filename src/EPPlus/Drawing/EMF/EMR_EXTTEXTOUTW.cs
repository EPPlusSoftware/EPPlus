using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Utils;
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
        internal byte[] Reference;
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

        public EMR_EXTTEXTOUTW(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            Bounds = br.ReadBytes(16);
            iGraphicsMode = br.ReadBytes(4);
            exScale = br.ReadBytes(4);
            eyScale = br.ReadBytes(4);
            Reference = br.ReadBytes(8);        //Signed, koordinat för var texten börjar. 
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

        byte[] CalculateDxSpacing()
        {
            var test = textSettings.GenericTextMeasurer;
            var aMesurement = (GenericFontMetricsTextMeasurer)textSettings.GenericTextMeasurer;
            aMesurement.MeasureTextInternal(stringBuffer, GenericFontMetricsTextMeasurerBase.GetKey(Font.elw.mFont.FontFamily, Font.elw.mFont.Style), Font.elw.mFont.Style, Font.elw.mFont.Size);
            var values = aMesurement.MeasureIndividualCharacters(stringBuffer, Font.elw.mFont, Ppi);

            int index = 0;
            foreach (uint val in values)
            {
                var bytes = BitConverter.GetBytes(val);
                bytes.CopyTo(DxBuffer, index);
                index += bytes.Length;
            }
            return DxBuffer;
        }
        public EMR_EXTTEXTOUTW(string Text)
        {
            Type = RECORD_TYPES.EMR_EXTTEXTOUTW;
            Bounds = new byte[16] { 0x13, 0x00, 0x00, 0x00, 0x24, 0x00, 0x00, 0x00, 0x4b, 0x00, 0x00, 0x00, 0x30, 0x00, 0x00, 0x00 };
            iGraphicsMode = new byte[4] { 0x02, 0x00, 0x00, 0x00 };
            exScale = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            eyScale = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            Reference = new byte[8] { 0x13, 0x00, 0x00, 0x00, 0x24, 0x00, 0x00, 0x00 };
            Options = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            Rectangle = new byte[16] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF };
            offString = 4 + 4 + 16 + 4 + 4 + 4 + 8 + 4 + 4 + 4 + 16 + 4;
            stringBuffer = Text;
            CalculateOffsets();
        }

        public EMR_EXTTEXTOUTW(string Text, int x, int y)
        {
            Type = RECORD_TYPES.EMR_EXTTEXTOUTW;
            Bounds = new byte[16] { 0x13, 0x00, 0x00, 0x00, 0x24, 0x00, 0x00, 0x00, 0x4b, 0x00, 0x00, 0x00, 0x30, 0x00, 0x00, 0x00 };
            iGraphicsMode = new byte[4] { 0x02, 0x00, 0x00, 0x00 };
            exScale = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            eyScale = new byte[4] { 0x00, 0x00, 0x00, 0x00 };

            var bx = BitConverter.GetBytes(x);
            var by = BitConverter.GetBytes(y);

            Reference = BinaryHelper.ConcatenateByteArrays(bx, by);

            //Reference = new byte[8] { 0x13, 0x00, 0x00, 0x00, 0x24, 0x00, 0x00, 0x00 };
            Options = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            Rectangle = new byte[16] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF };
            offString = 4 + 4 + 16 + 4 + 4 + 4 + 8 + 4 + 4 + 4 + 16 + 4;
            stringBuffer = Text;
            CalculateOffsets();
        }

        private void CalculateOffsets()
        {
            Chars = (uint)stringBuffer.Length;
            offDx = offString + (uint)stringBuffer.Length * 2;

            padding = (int)offDx;
            offDx += 4 - (offDx % 4);
            padding = (int)(offDx) - padding;

            byte[] prevBuff = new byte[DxBuffer.Length];
            DxBuffer.CopyTo(prevBuff, 0);

            DxBuffer = new byte[stringBuffer.Length * 4];
            CalculateDxSpacing();
            Size = offDx + (uint)DxBuffer.Length;
        }

        public override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(Bounds);
            bw.Write(iGraphicsMode);
            bw.Write(exScale);
            bw.Write(eyScale);
            bw.Write(Reference);
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
