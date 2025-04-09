using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
namespace OfficeOpenXml.Style.HeaderFooterTextFormat
{
    public class HFTextCollection : IEnumerable<HFText>
    {
        const string ARG_TO_LONG_EXCEPTION_TEXT = "Header and Footer texts cannot exceed 255 characters.";

        private List<HFText> _textCollection = new List<HFText>();
        private readonly int _defaultFontSize;
        private ImageInfo _imageInfo = null;

        internal readonly PictureAlignment alignment;
        internal ExcelWorksheet _ws;
        internal HFTextCollection lane1;
        internal HFTextCollection lane2;
        internal ExcelHeaderFooterText headerFooter;


        public ExcelVmlDrawingPicture VmlDrawingPicutre { get; set; } = null;


        public static int TextLength { get; private set; } = 0;


        public HFText this[int index]
        {
            get
            {
                return _textCollection[index];
            }
            set
            {
                _textCollection[index] = value;
            }
        }

        internal HFTextCollection(ExcelWorksheet ws, ExcelHeaderFooterText headerFooter, PictureAlignment alignment, int defaultFontSize)
        {
            _ws = ws;
            this.alignment = alignment;
            Clear();
            _defaultFontSize = defaultFontSize;
            this.headerFooter = headerFooter;
        }

        internal void ValidateTextLength()
        {
            string text1 = WriteHeaderFooterFormat();
            string text2 = lane1 == null ? string.Empty : lane1.WriteHeaderFooterFormat();
            string text3 = lane2 == null ? string.Empty : lane2.WriteHeaderFooterFormat();
            int length = text1.Length + text2.Length + text3.Length - 2;//-2 here because it appears Excel does not count the other lanes format codes (&R, &C, &L).
            if (length > 255)
            {
                throw new ArgumentOutOfRangeException(ARG_TO_LONG_EXCEPTION_TEXT);
            }
            TextLength = length;
        }

        public void Add(HFText textItem)
        {
            _textCollection.Add(textItem);
            ValidateTextLength();
        }
        public HFText AddText(string text)
        {
            HFText hfText = new HFText(text);
            _textCollection.Add(hfText);
            ValidateTextLength();
            return hfText;
        }
        public HFText AddPageNumber(int number = 0)
        {
            HFText hft = new HFText();
            hft.FormatCode = HFFormattingCodes.PageNumber;
            if (number > 0 || number < 0)
            {
                hft.PageNumberSuffix = number >= 0 ? "+" : "-";
                hft.PageNumberSuffix += Math.Abs(number);
            }
            Add(hft);
            return hft;
        }
        public HFText AddNumberOfPages()
        {
            return AddFormatCode(HFFormattingCodes.NumberOfPages);
        }
        public HFText AddSheetName()
        {
            return AddFormatCode(HFFormattingCodes.SheetName);
        }
        public HFText AddFilePath()
        {
            return AddFormatCode(HFFormattingCodes.FilePath);
        }
        public HFText AddFileName()
        {
            return AddFormatCode(HFFormattingCodes.FileName);
        }
        public HFText AddCurrentDate()
        {
            return AddFormatCode(HFFormattingCodes.CurrentDate);
        }
        public HFText AddCurrentTime()
        {
            return AddFormatCode(HFFormattingCodes.CurrentTime);
        }
        public HFText AddImage(FileInfo pictureFile)
        {
            headerFooter.InsertPicture(pictureFile, (PictureAlignment)alignment);
            return AddFormatCode(HFFormattingCodes.Image);
        }
        public HFText AddImage (Stream pictureStream, ePictureType pictureType)
        {
            headerFooter.InsertPicture(pictureStream, pictureType, (PictureAlignment)alignment);
            return AddFormatCode(HFFormattingCodes.Image);
        }
        private HFText AddFormatCode(HFFormattingCodes formatCode)
        {
            HFText hft = new HFText();
            hft.FormatCode = formatCode;
            Add(hft);
            return hft;
        }

        public void Insert(int index, HFText item)
        {
            if (index < 0 || index >= _textCollection.Count) throw new IndexOutOfRangeException("index was out of range.");
            if (item == null) throw new ArgumentException("Item can't be null", "item");
            _textCollection.Insert(index, item);
            ValidateTextLength();
        }
        public HFText InsertText(int index, string text)
        {
            if (index < 0 || index >= _textCollection.Count) throw new IndexOutOfRangeException("index was out of range.");
            if (text == null) throw new ArgumentException("Text can't be null", "text");
            HFText hFText = new HFText(text);
            _textCollection.Insert(index, hFText);
            ValidateTextLength();
            return hFText;
        }
        //Insert pagenumber, date...
        public HFText InsertPageNumber(int index, int number = 0)
        {
            HFText hft = new HFText();
            hft.FormatCode = HFFormattingCodes.PageNumber;
            if (number > 0 || number < 0)
            {
                hft.PageNumberSuffix = number >= 0 ? "+" : "-";
                hft.PageNumberSuffix += Math.Abs(number);
            }
            Insert(index, hft);
            return hft;
        }
        public HFText InsertNumberOfPages(int index)
        {
            return InsertFormatCode(index, HFFormattingCodes.NumberOfPages);
        }
        public HFText InsertSheetName(int index)
        {
            return InsertFormatCode(index, HFFormattingCodes.SheetName);
        }
        public HFText InsertFilePath(int index)
        {
            return InsertFormatCode(index, HFFormattingCodes.FilePath);
        }
        public HFText InsertFileName(int index)
        {
            return InsertFormatCode(index, HFFormattingCodes.FileName);
        }
        public HFText InsertCurrentDate(int index)
        {
            return InsertFormatCode(index, HFFormattingCodes.CurrentDate);
        }
        public HFText InsertCurrentTime(int index)
        {
            return InsertFormatCode(index, HFFormattingCodes.CurrentTime);
        }
        public HFText InsertImage(int index, FileInfo pictureFile)
        {
            string id = headerFooter.ValidateImage((PictureAlignment)alignment);
            if (!pictureFile.Exists)
            {
                throw (new FileNotFoundException(string.Format("{0} is missing", pictureFile.FullName)));
            }
            var uriPic = XmlHelper.GetNewUri(_ws._package.ZipPackage, "/xl/media/" + pictureFile.Name.Substring(0, pictureFile.Name.Length - pictureFile.Extension.Length) + "{0}" + pictureFile.Extension);
            var imgBytes = File.ReadAllBytes(pictureFile.FullName);
            var ii = _ws.Workbook._package.PictureStore.AddImage(imgBytes, uriPic, null);
            var hft = InsertFormatCode(index, HFFormattingCodes.Image);
            VmlDrawingPicutre = headerFooter.AddImage(id, ii);
            return hft;
        }
        public HFText InsertImage(int index, Stream pictureStream, ePictureType pictureType)
        {
            string id = headerFooter.ValidateImage((PictureAlignment)alignment);
            var imgBytes = new byte[pictureStream.Length];
            pictureStream.Seek(0, SeekOrigin.Begin);
            var r = pictureStream.Read(imgBytes, 0, imgBytes.Length);
            _imageInfo = _ws.Workbook._package.PictureStore.AddImage(imgBytes, null, pictureType);
            var hft = InsertFormatCode(index, HFFormattingCodes.Image);
            VmlDrawingPicutre = headerFooter.AddImage(id, _imageInfo);
            return hft;
        }
        private HFText InsertFormatCode(int index, HFFormattingCodes formatCode)
        {
            if (index < 0 || index >= _textCollection.Count) throw new IndexOutOfRangeException("index was out of range.");
            if (formatCode == HFFormattingCodes.Text) throw new ArgumentException("HFFormattingCode cannot be Text", "formatcode");
            HFText hFText = new HFText();
            hFText.FormatCode = formatCode;
            _textCollection.Insert(index, hFText);
            ValidateTextLength();
            return hFText;
        }

        public void Remove(HFText item)
        {
            if(item.FormatCode == HFFormattingCodes.Image)
            {
                RemoveImage();
                return;
            }
            _textCollection.Remove(item);
            ValidateTextLength();
        }

        public void RemoveAt(int index)
        {
            if (_textCollection[index].FormatCode == HFFormattingCodes.Image)
            {
                RemoveImage();
                return;
            }
            _textCollection.RemoveAt(index);
            ValidateTextLength();
        }

        public void RemoveImage()
        {
            headerFooter.RemoveImage(VmlDrawingPicutre);
            VmlDrawingPicutre = null;
            _ws.Workbook._package.PictureStore.RemoveReference(_imageInfo.Uri);
            foreach (HFText hft in _textCollection)
            {
                if (hft.FormatCode == HFFormattingCodes.Image)
                {
                    _textCollection.Remove(hft);
                    break;
                }
            }
        }

        public void Clear()
        {
            _textCollection.Clear();
            switch (alignment)
            {
                case PictureAlignment.Left:
                    Add(new HFText("&L"));
                    break;
                case PictureAlignment.Centered:
                    Add(new HFText("&C"));
                    break;
                case PictureAlignment.Right:
                    Add(new HFText("&R"));
                    break;
            }
        }
        public int Count
        {
            get { return _textCollection.Count; }
        }

        public string Text
        {
            get
            {
                StringBuilder sb = new StringBuilder();
                foreach (var item in _textCollection)
                {
                    sb.Append(item.Text);
                }
                return sb.ToString();
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    Clear();
                }
                else if (Count == 0)
                {
                    AddText(value);
                }
                else if (Count > 1)
                {
                    while (Count != 1)
                    {
                        RemoveAt(1);
                    }
                    this[0].Text = value;
                }
                else if (Count == 1)
                {
                    this[0].Text = value;
                }
            }
        }

        private Color GetColor(string c, ref eThemeSchemeColor? theme, ref double? tint)
        {
            if (c.Contains('-') || c.Contains('+'))
            {
                theme = (eThemeSchemeColor)int.Parse(""+c[0] + c[1]);
                var themeColor = _ws.Workbook.ThemeManager.GetOrCreateTheme().ColorScheme.GetColorByEnum((eThemeSchemeColor)theme);
                var color = Utils.ColorConverter.GetThemeColor(themeColor);
                tint = double.Parse((""+c[3] + c[4] + c[5]));
                tint = c[2] == '-' ? -(tint/100) : tint / 100;
                return Utils.ColorConverter.ApplyTint(color, (double)tint);
            }
            else
            {
                var argb = int.Parse(c, System.Globalization.NumberStyles.HexNumber);
                return Color.FromArgb((argb >> 24) & 0xFF, (argb >> 16) & 0xFF, (argb >> 8) & 0xFF, argb & 0xFF);
            }
        }

        internal void ReadHeaderFooterFormat(string hfText)
        {
            HFText temp = new HFText();
            bool quoteTag = false;
            bool stop = false;
            bool writeFormatCode = false;
            int i = 0;
            while (i < hfText.Length)
            {
                if (hfText[i] == '&' && !writeFormatCode)
                {
                    i++;
                    switch (hfText[i])
                    {
                        //Text Field
                        case 'L':
                            if (alignment != PictureAlignment.Left) stop = true;
                            break;
                        case 'C':
                            if (alignment != PictureAlignment.Centered) stop = true;
                            break;
                        case 'R':
                            if (alignment != PictureAlignment.Right) stop = true;
                            break;
                        //Constants
                        case 'P':
                            temp.FormatCode = HFFormattingCodes.PageNumber;
                            writeFormatCode = true;
                            break;
                        case 'N':
                            temp.FormatCode = HFFormattingCodes.NumberOfPages;
                            writeFormatCode = true;
                            break;
                        case 'A':
                            temp.FormatCode = HFFormattingCodes.SheetName;
                            writeFormatCode = true;
                            break;
                        case 'Z':
                            temp.FormatCode = HFFormattingCodes.FilePath;
                            writeFormatCode = true;
                            break;
                        case 'F':
                            temp.FormatCode = HFFormattingCodes.FileName;
                            writeFormatCode = true;
                            break;
                        case 'D':
                            temp.FormatCode = HFFormattingCodes.CurrentDate;
                            writeFormatCode = true;
                            break;
                        case 'T':
                            temp.FormatCode = HFFormattingCodes.CurrentTime;
                            writeFormatCode = true;
                            break;
                        case 'G':
                            temp.FormatCode = HFFormattingCodes.Image;
                            foreach(ExcelVmlDrawingPicture v in _ws.HeaderFooter.Pictures)
                            {
                                if(v.Id == alignment.ToString()[0] + headerFooter.HeaderFooterAlignment)
                                {
                                    VmlDrawingPicutre = v;
                                    _imageInfo = _ws.Workbook._package.PictureStore.LoadImage(v.Image.ImageBytes, v.ImageUri, _ws._package.ZipPackage.GetPart(v.ImageUri)); //remove this line or something when picstore works
                                }
                            }
                            writeFormatCode = true;
                            break;
                        //Text Formating
                        case 'B':
                            temp.Bold = !temp.Bold;
                            break;
                        case 'I':
                            temp.Italic = !temp.Italic;
                            break;
                        case 'U':
                            temp.Underline = !temp.Underline;
                            break;
                        case 'E':
                            temp.DoubleUnderline = !temp.DoubleUnderline;
                            break;
                        case 'H':
                            temp.Shadow = !temp.Shadow;
                            break;
                        case 'K':
                            string color = hfText.Substring(i + 1, 6);
                            eThemeSchemeColor? theme = null;
                            double? tint = null;
                            temp.Color = GetColor(color, ref theme, ref tint );
                            temp.Theme = theme;
                            temp.Tint = tint;
                            i += 6;
                            break;
                        case 'O':
                            temp.Outline = !temp.Outline;
                            break;
                        case 'S':
                            temp.Striketrough = !temp.Striketrough;
                            break;
                        case 'X':
                            temp.SuperScript = !temp.SuperScript;
                            break;
                        case 'Y':
                            temp.SubScript = !temp.SubScript;
                            break;
                        case '&':
                            temp.Text += "&";
                            i++;
                            continue;
                        case '"':
                            quoteTag = true;
                            int end = hfText.IndexOf('"', i+1);
                            if(end > i)
                            {
                                string tag = hfText.Substring(i + 1, end - i - 1);
                                string[] parts = tag.Split(',');
                                if(parts.Length > 0 && parts[0] != "-")
                                {
                                    temp.FontName = parts[0];
                                }
                                else if (parts[0] == "-")
                                {
                                    temp.FontName = "-";
                                }
                                for(int j = 1; j< parts.Length; j++)
                                {
                                    string trim = parts[j].Trim().ToLower();
                                    if(trim.Contains("bold")) temp.Bold = true;
                                    if(trim.Contains("italic")) temp.Italic = true;
                                    if (trim.Contains("regular"))
                                    {
                                        temp.Bold = false;
                                        temp.Italic = false;
                                    }
                                }
                                i = end;
                            }
                            break;
                        default:
                            //Font Size
                            if (char.IsDigit(hfText[i]))
                            {
                                int num = i;
                                while (num < hfText.Length && char.IsDigit(hfText[num]))
                                    num++;
                                temp.FontSize = int.Parse(hfText.Substring(i, num - i));
                                i = num - 1;
                            }
                            break;
                    }
                    i++;
                }
                else
                {
                    if (!writeFormatCode)
                    {
                        while (true)
                        {
                            temp.Text += hfText[i];
                            i++;
                            if (i >= hfText.Length)
                            {
                                break;
                            }
                            else if (hfText[i] == '&' && hfText[i + 1] == '&')
                            {
                                i++;
                            }
                            else if (hfText[i] == '&' && hfText[i + 1] != '&')
                            {
                                break;
                            }
                        }
                    }
                    else if(writeFormatCode && temp.FormatCode == HFFormattingCodes.PageNumber)
                    {
                        if (hfText[i] == '+' || hfText[i] == '-')
                        {
                            temp.PageNumberSuffix += hfText[i];
                        }
                        i++;
                        while (char.IsDigit( hfText[i]))
                        {
                            temp.PageNumberSuffix += hfText[i];
                            i++;
                        }
                    }
                    HFText newText = new HFText
                    {
                        Text = temp.Text,
                        Bold = temp.Bold,
                        Italic = temp.Italic,
                        Underline = temp.Underline,
                        DoubleUnderline = temp.DoubleUnderline,
                        Shadow = temp.Shadow,
                        Color = temp.Color,
                        Theme = temp.Theme,
                        Tint = temp.Tint,
                        Outline = temp.Outline,
                        Striketrough = temp.Striketrough,
                        SuperScript = temp.SuperScript,
                        SubScript = temp.SubScript,
                        FontSize = temp.FontSize,
                        FontName = temp.FontName,
                        FormatCode = temp.FormatCode,
                        PageNumberSuffix = temp.PageNumberSuffix,
                    };
                    _textCollection.Add(newText);
                    writeFormatCode = false;
                    temp.FormatCode = HFFormattingCodes.Text;
                    temp.Text = "";
                    if (quoteTag)
                    {
                        temp = new HFText();
                        quoteTag = false;
                    }
                }
                if (stop) break;
            }
            if (writeFormatCode)
            {
                HFText newText = new HFText
                {
                    Text = temp.Text,
                    Bold = temp.Bold,
                    Italic = temp.Italic,
                    Underline = temp.Underline,
                    DoubleUnderline = temp.DoubleUnderline,
                    Shadow = temp.Shadow,
                    Color = temp.Color,
                    Theme = temp.Theme,
                    Tint = temp.Tint,
                    Outline = temp.Outline,
                    Striketrough = temp.Striketrough,
                    SuperScript = temp.SuperScript,
                    SubScript = temp.SubScript,
                    FontSize = temp.FontSize,
                    FontName = temp.FontName,
                    FormatCode = temp.FormatCode,
                    PageNumberSuffix = temp.PageNumberSuffix,
                };
                _textCollection.Add(newText);
            }
        }

        internal string WriteHeaderFooterFormat()
        {
            string hfstring = _textCollection[0].Text;
            for(int i = 1; i < _textCollection.Count; i++)
            {
                hfstring += WriteHeaderFooter2(_textCollection[i], _textCollection[i - 1]);
            }
            return hfstring;
        }

        private string WriteHeaderFooter2(HFText current, HFText prev)
        {
            string hfstring = "";

            if (!(current.FontName == prev.FontName && current.Bold == prev.Bold && current.Italic == prev.Italic))
            {
                var fontParts = new List<string> { current.FontName ?? "-" };

                if (current.Bold && current.Italic) fontParts.Add("Bold Italic");
                else if (current.Bold) fontParts.Add("Bold");
                else if (current.Italic) fontParts.Add("Italic");

                if (!current.Bold && !current.Italic)
                    fontParts.Add("Regular");

                hfstring += $"&\"{string.Join(",", fontParts.ToArray())}\"";
            }

            if (current.FontSize != null && prev.FontSize != null)
            {
                if (current.FontSize != prev.FontSize)
                {
                    hfstring += "&" + current.FontSize;
                    if (char.IsDigit(current.Text[0]))
                    {
                        hfstring += " ";
                    }
                }
            }
            else if (current.FontSize != null && prev.FontSize == null)
            {
                hfstring += "&" + current.FontSize;
                if (char.IsDigit(current.Text[0]))
                {
                    hfstring += " ";
                }
            } else if (current.FontSize == null && prev.FontSize != null )
            {
                hfstring += "&" + _defaultFontSize;
            }

            if (current.Underline && !prev.Underline) hfstring += "&U";
            else if(!current.Underline && prev.Underline) hfstring += "&U";

            if (current.DoubleUnderline && !prev.DoubleUnderline) hfstring += "&E";
            else if(!current.DoubleUnderline && prev.DoubleUnderline) hfstring += "&E";


            if (current.Shadow) hfstring += "&H";

            string currentColor = current.GetThemeOrColorAsString();
            string prevColor = prev.GetThemeOrColorAsString();
            if (currentColor != prevColor) hfstring += currentColor;

            if (current.Outline) hfstring += "&O";

            if (current.Striketrough) hfstring += "&S";

            if (current.SuperScript) hfstring += "&X";

            if (current.SubScript) hfstring += "&Y";

            switch (current.FormatCode)
            {
                case HFFormattingCodes.PageNumber:
                    hfstring += "&P" + current.PageNumberSuffix;
                    break;
                case HFFormattingCodes.NumberOfPages:
                    hfstring += "&N";
                    break;
                case HFFormattingCodes.SheetName:
                    hfstring += "&A";
                    break;
                case HFFormattingCodes.FilePath:
                    hfstring += "&Z";
                    break;
                case HFFormattingCodes.FileName:
                    hfstring += "&F";
                    break;
                case HFFormattingCodes.CurrentDate:
                    hfstring += "&D";
                    break;
                case HFFormattingCodes.CurrentTime:
                    hfstring += "&T";
                    break;
                case HFFormattingCodes.Image:
                    hfstring += "&G";
                    if (current.Text == "&") hfstring += "&";
                    hfstring += current.Text;
                    break;
                case HFFormattingCodes.Text:
                    if (current.Text == "&") hfstring += "&";
                    hfstring += current.Text;
                    break;
            }
            return hfstring;
        }


        IEnumerator IEnumerable.GetEnumerator()
        {
            return _textCollection.GetEnumerator();
        }

        IEnumerator<HFText> IEnumerable<HFText>.GetEnumerator()
        {
            return _textCollection.GetEnumerator();
        }
    }
}
