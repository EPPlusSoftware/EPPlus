using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
namespace OfficeOpenXml.Style.HeaderFooterTextFormat
{
    /// <summary>
    /// The collection of ExcelHeaderFooterTextItems.
    /// </summary>
    public class ExcelHeaderFooterTextCollection : IEnumerable<ExcelHeaderFooterTextItem>
    {
        internal const string ARG_TO_LONG_EXCEPTION_TEXT = "Header and Footer texts cannot exceed 255 characters.";

        private List<ExcelHeaderFooterTextItem> _textCollection = new List<ExcelHeaderFooterTextItem>();
        private readonly int _defaultFontSize;
        private ImageInfo _imageInfo = null;

        internal readonly PictureAlignment alignment;
        internal ExcelWorksheet _ws;
        internal ExcelHeaderFooterTextCollection lane1;
        internal ExcelHeaderFooterTextCollection lane2;
        internal ExcelHeaderFooterText headerFooter;

        /// <summary>
        /// The drawing vml reference to inserted picture.
        /// </summary>
        public ExcelVmlDrawingPicture Picture { get; set; } = null;

        /// <summary>
        /// The length of the text of left, center and right header footer fields.
        /// </summary>
        public int TextLength { get; internal set; } = 0;

        /// <summary>
        /// Returns the ExcelHeaderFooterTextItem at index
        /// </summary>
        /// <param name="index">the index of item to return.</param>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem this[int index]
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

        internal ExcelHeaderFooterTextCollection(ExcelWorksheet ws, ExcelHeaderFooterText headerFooter, PictureAlignment alignment, int defaultFontSize)
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
            if (lane1 != null)
            {
                lane1.TextLength = length;
            }
            if (lane2 != null)
            {
                lane2.TextLength = length;
            }
            TextLength = length;
        }

        /// <summary>
        /// Add an ExcelHeaderFooterTextItem to the end of the collection.
        /// </summary>
        /// <param name="textItem"></param>
        public void Add(ExcelHeaderFooterTextItem textItem)
        {
            _textCollection.Add(textItem);
            ValidateTextLength();
        }
        /// <summary>
        /// Add a string to the end of the collection.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem AddText(string text)
        {
            ExcelHeaderFooterTextItem hfText = new ExcelHeaderFooterTextItem(text);
            _textCollection.Add(hfText);
            ValidateTextLength();
            return hfText;
        }
        /// <summary>
        /// Add the page number to the end of the collection.
        /// </summary>
        /// <param name="number">Amount to offset the page number with. Default is 0.</param>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem AddPageNumber(int number = 0)
        {
            ExcelHeaderFooterTextItem hft = new ExcelHeaderFooterTextItem();
            hft.FormatCode = ExcelHeaderFooterFormattingCodes.PageNumber;
            if (number > 0 || number < 0)
            {
                hft.PageNumberSuffix = number >= 0 ? "+" : "-";
                hft.PageNumberSuffix += Math.Abs(number);
            }
            Add(hft);
            return hft;
        }
        /// <summary>
        /// Add the number of pages to the end of the collection.
        /// </summary>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem AddNumberOfPages()
        {
            return AddFormatCode(ExcelHeaderFooterFormattingCodes.NumberOfPages);
        }
        /// <summary>
        /// Add the sheet name to the end of the collection.
        /// </summary>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem AddSheetName()
        {
            return AddFormatCode(ExcelHeaderFooterFormattingCodes.SheetName);
        }
        /// <summary>
        /// Add the file path to the end of the collection.
        /// </summary>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem AddFilePath()
        {
            return AddFormatCode(ExcelHeaderFooterFormattingCodes.FilePath);
        }
        /// <summary>
        /// Add the file name to the end of the collection.
        /// </summary>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem AddFileName()
        {
            return AddFormatCode(ExcelHeaderFooterFormattingCodes.FileName);
        }
        /// <summary>
        /// Add the current date to the end of the collection.
        /// </summary>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem AddCurrentDate()
        {
            return AddFormatCode(ExcelHeaderFooterFormattingCodes.CurrentDate);
        }
        /// <summary>
        /// Add the current time to the end of the collection.
        /// </summary>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem AddCurrentTime()
        {
            return AddFormatCode(ExcelHeaderFooterFormattingCodes.CurrentTime);
        }
        /// <summary>
        /// Add a picture to the end of the collection.
        /// </summary>
        /// <param name="pictureFile">The FileInfo object for the picture.</param>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem AddImage(FileInfo pictureFile)
        {
            headerFooter.InsertPicture(pictureFile, (PictureAlignment)alignment);
            return AddFormatCode(ExcelHeaderFooterFormattingCodes.Image);
        }
        /// <summary>
        /// Add the file path to the end of the collection using a stream.
        /// </summary>
        /// <param name="pictureStream">The picture stream.</param>
        /// <param name="pictureType">The file type of the picture.</param>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem AddImage (Stream pictureStream, ePictureType pictureType)
        {
            headerFooter.InsertPicture(pictureStream, pictureType, (PictureAlignment)alignment);
            return AddFormatCode(ExcelHeaderFooterFormattingCodes.Image);
        }
        private ExcelHeaderFooterTextItem AddFormatCode(ExcelHeaderFooterFormattingCodes formatCode)
        {
            ExcelHeaderFooterTextItem hft = new ExcelHeaderFooterTextItem();
            hft.FormatCode = formatCode;
            Add(hft);
            return hft;
        }

        /// <summary>
        /// Insert an ExcelHeaderFooterTextItem at specified position in the collection.
        /// </summary>
        /// <param name="index"></param>
        /// <param name="item"></param>
        /// <exception cref="IndexOutOfRangeException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public void Insert(int index, ExcelHeaderFooterTextItem item)
        {
            if (index < 0 || index >= _textCollection.Count) throw new IndexOutOfRangeException("index was out of range.");
            if (item == null) throw new ArgumentException("Item can't be null", "item");
            _textCollection.Insert(index, item);
            ValidateTextLength();
        }
        /// <summary>
        /// Insert a string at specified position in the collection.
        /// </summary>
        /// <param name="index"></param>
        /// <param name="text"></param>
        /// <returns></returns>
        /// <exception cref="IndexOutOfRangeException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public ExcelHeaderFooterTextItem InsertText(int index, string text)
        {
            if (index < 0 || index >= _textCollection.Count) throw new IndexOutOfRangeException("index was out of range.");
            if (text == null) throw new ArgumentException("Text can't be null", "text");
            ExcelHeaderFooterTextItem hFText = new ExcelHeaderFooterTextItem(text);
            _textCollection.Insert(index, hFText);
            ValidateTextLength();
            return hFText;
        }
        /// <summary>
        /// Insert the page number at specified position in the collection.
        /// </summary>
        /// <param name="index"></param>
        /// <param name="number"></param>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem InsertPageNumber(int index, int number = 0)
        {
            ExcelHeaderFooterTextItem hft = new ExcelHeaderFooterTextItem();
            hft.FormatCode = ExcelHeaderFooterFormattingCodes.PageNumber;
            if (number > 0 || number < 0)
            {
                hft.PageNumberSuffix = number >= 0 ? "+" : "-";
                hft.PageNumberSuffix += Math.Abs(number);
            }
            Insert(index, hft);
            return hft;
        }
        /// <summary>
        /// Insert the number of pages at specified position in the collection.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem InsertNumberOfPages(int index)
        {
            return InsertFormatCode(index, ExcelHeaderFooterFormattingCodes.NumberOfPages);
        }
        /// <summary>
        /// Insert the sheet name at specified position in the collection.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem InsertSheetName(int index)
        {
            return InsertFormatCode(index, ExcelHeaderFooterFormattingCodes.SheetName);
        }
        /// <summary>
        /// Insert the file path at specified position in the collection.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem InsertFilePath(int index)
        {
            return InsertFormatCode(index, ExcelHeaderFooterFormattingCodes.FilePath);
        }
        /// <summary>
        /// Insert the file name at specified position in the collection.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem InsertFileName(int index)
        {
            return InsertFormatCode(index, ExcelHeaderFooterFormattingCodes.FileName);
        }
        /// <summary>
        /// Insert the current date at specified position in the collection.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem InsertCurrentDate(int index)
        {
            return InsertFormatCode(index, ExcelHeaderFooterFormattingCodes.CurrentDate);
        }
        /// <summary>
        /// Insert the current time at specified position in the collection.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem InsertCurrentTime(int index)
        {
            return InsertFormatCode(index, ExcelHeaderFooterFormattingCodes.CurrentTime);
        }
        /// <summary>
        /// Insert a picture at specified position in the collection.
        /// </summary>
        /// <param name="index"></param>
        /// <param name="pictureFile"></param>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem InsertImage(int index, FileInfo pictureFile)
        {
            string id = headerFooter.ValidateImage((PictureAlignment)alignment);
            if (!pictureFile.Exists)
            {
                throw (new FileNotFoundException(string.Format("{0} is missing", pictureFile.FullName)));
            }
            var uriPic = XmlHelper.GetNewUri(_ws._package.ZipPackage, "/xl/media/" + pictureFile.Name.Substring(0, pictureFile.Name.Length - pictureFile.Extension.Length) + "{0}" + pictureFile.Extension);
            var imgBytes = File.ReadAllBytes(pictureFile.FullName);
            var ii = _ws.Workbook._package.PictureStore.AddImage(imgBytes, uriPic, null);
            var hft = InsertFormatCode(index, ExcelHeaderFooterFormattingCodes.Image);
            Picture = headerFooter.AddImage(id, ii);
            return hft;
        }
        /// <summary>
        /// Insert a picture at specified position in the collection using a stream.
        /// </summary>
        /// <param name="index"></param>
        /// <param name="pictureStream"></param>
        /// <param name="pictureType"></param>
        /// <returns></returns>
        public ExcelHeaderFooterTextItem InsertImage(int index, Stream pictureStream, ePictureType pictureType)
        {
            string id = headerFooter.ValidateImage((PictureAlignment)alignment);
            var imgBytes = new byte[pictureStream.Length];
            pictureStream.Seek(0, SeekOrigin.Begin);
            var r = pictureStream.Read(imgBytes, 0, imgBytes.Length);
            _imageInfo = _ws.Workbook._package.PictureStore.AddImage(imgBytes, null, pictureType);
            var hft = InsertFormatCode(index, ExcelHeaderFooterFormattingCodes.Image);
            Picture = headerFooter.AddImage(id, _imageInfo);
            return hft;
        }
        private ExcelHeaderFooterTextItem InsertFormatCode(int index, ExcelHeaderFooterFormattingCodes formatCode)
        {
            if (index < 0 || index >= _textCollection.Count) throw new IndexOutOfRangeException("index was out of range.");
            if (formatCode == ExcelHeaderFooterFormattingCodes.Text) throw new ArgumentException("HFFormattingCode cannot be Text", "formatcode");
            ExcelHeaderFooterTextItem hFText = new ExcelHeaderFooterTextItem();
            hFText.FormatCode = formatCode;
            _textCollection.Insert(index, hFText);
            ValidateTextLength();
            return hFText;
        }

        /// <summary>
        /// Remove the specified item from the collection.
        /// </summary>
        /// <param name="item"></param>
        public void Remove(ExcelHeaderFooterTextItem item)
        {
            if(item.FormatCode == ExcelHeaderFooterFormattingCodes.Image)
            {
                RemovePicture();
                return;
            }
            _textCollection.Remove(item);
            ValidateTextLength();
        }
        /// <summary>
        /// Remove the specified item at index in collection.
        /// </summary>
        /// <param name="index"></param>
        public void RemoveAt(int index)
        {
            if (_textCollection[index].FormatCode == ExcelHeaderFooterFormattingCodes.Image)
            {
                RemovePicture();
                return;
            }
            _textCollection.RemoveAt(index);
            ValidateTextLength();
        }
        /// <summary>
        /// Remove the picture from the header or footer.
        /// </summary>
        public void RemovePicture()
        {
            headerFooter.RemoveImage(Picture);
            Picture = null;
            //_ws.Workbook._package.PictureStore.RemoveReference(_imageInfo.Uri);
            foreach (ExcelHeaderFooterTextItem hft in _textCollection)
            {
                if (hft.FormatCode == ExcelHeaderFooterFormattingCodes.Image)
                {
                    _textCollection.Remove(hft);
                    break;
                }
            }
        }
        /// <summary>
        /// Clear the collection.
        /// </summary>
        public void Clear()
        {
            _textCollection.Clear();
            switch (alignment)
            {
                case PictureAlignment.Left:
                    Add(new ExcelHeaderFooterTextItem("&L"));
                    break;
                case PictureAlignment.Centered:
                    Add(new ExcelHeaderFooterTextItem("&C"));
                    break;
                case PictureAlignment.Right:
                    Add(new ExcelHeaderFooterTextItem("&R"));
                    break;
            }
        }
        /// <summary>
        /// The number of objects in the collection.
        /// </summary>
        public int Count
        {
            get { return _textCollection.Count; }
        }
        /// <summary>
        /// Get the raw text of the collection.
        /// </summary>
        public string Text
        {
            get
            {
                return WriteHeaderFooterFormat();
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
                ValidateTextLength();
            }
        }

        private Color GetColor(string c, ref eThemeSchemeColor? theme, ref double? tint)
        {
            if (c.Contains('-') || c.Contains('+'))
            {
                theme = (eThemeSchemeColor)int.Parse(""+c[0] + c[1]);
                var themeColor = _ws.Workbook.ThemeManager.GetOrCreateTheme().ColorScheme.GetColorByEnum((eThemeSchemeColor)theme);
                var color = Utils.TypeConversion.ColorConverter.GetThemeColor(themeColor);
                tint = double.Parse((""+c[3] + c[4] + c[5]));
                tint = c[2] == '-' ? -(tint/100) : tint / 100;
                return Utils.TypeConversion.ColorConverter.ApplyTint(color, (double)tint);
            }
            else
            {
                var argb = int.Parse(c, System.Globalization.NumberStyles.HexNumber);
                return Color.FromArgb((argb >> 24) & 0xFF, (argb >> 16) & 0xFF, (argb >> 8) & 0xFF, argb & 0xFF);
            }
        }

        internal void ReadHeaderFooterFormat(string hfText)
        {
            ExcelHeaderFooterTextItem temp = new ExcelHeaderFooterTextItem();
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
                            temp.FormatCode = ExcelHeaderFooterFormattingCodes.PageNumber;
                            writeFormatCode = true;
                            break;
                        case 'N':
                            temp.FormatCode = ExcelHeaderFooterFormattingCodes.NumberOfPages;
                            writeFormatCode = true;
                            break;
                        case 'A':
                            temp.FormatCode = ExcelHeaderFooterFormattingCodes.SheetName;
                            writeFormatCode = true;
                            break;
                        case 'Z':
                            temp.FormatCode = ExcelHeaderFooterFormattingCodes.FilePath;
                            writeFormatCode = true;
                            break;
                        case 'F':
                            temp.FormatCode = ExcelHeaderFooterFormattingCodes.FileName;
                            writeFormatCode = true;
                            break;
                        case 'D':
                            temp.FormatCode = ExcelHeaderFooterFormattingCodes.CurrentDate;
                            writeFormatCode = true;
                            break;
                        case 'T':
                            temp.FormatCode = ExcelHeaderFooterFormattingCodes.CurrentTime;
                            writeFormatCode = true;
                            break;
                        case 'G':
                            temp.FormatCode = ExcelHeaderFooterFormattingCodes.Image;
                            foreach(ExcelVmlDrawingPicture v in _ws.HeaderFooter.Pictures)
                            {
                                if(v.Id == alignment.ToString()[0] + headerFooter.HeaderFooterAlignment)
                                {
                                    Picture = v;
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
                            int end = hfText.IndexOf('"', i+1);
                            if(end > i)
                            {
                                string tag = hfText.Substring(i + 1, end - i - 1);
                                string[] parts = tag.Split(',');
                                if(parts.Length > 0)
                                {
                                    temp.FontName = parts[0] == "-" ? string.Empty : parts[0];
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
                    else if(writeFormatCode && temp.FormatCode == ExcelHeaderFooterFormattingCodes.PageNumber)
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
                    ExcelHeaderFooterTextItem newText = new ExcelHeaderFooterTextItem
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
                    temp.FormatCode = ExcelHeaderFooterFormattingCodes.Text;
                    temp.Text = "";
                }
                if (stop) break;
            }
            if (writeFormatCode)
            {
                ExcelHeaderFooterTextItem newText = new ExcelHeaderFooterTextItem
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
                hfstring += ParseToHeaderFooterFormat(_textCollection[i], _textCollection[i - 1]);
            }
            return hfstring;
        }

        private string ParseToHeaderFooterFormat(ExcelHeaderFooterTextItem current, ExcelHeaderFooterTextItem prev)
        {
            string hfstring = "";

            if (!(current.FontName == prev.FontName && current.Bold == prev.Bold && current.Italic == prev.Italic))
            {
                var fontParts = new List<string> { string.IsNullOrEmpty(current.FontName) ? "-" : current.FontName };

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
                    if (current.Text != null && char.IsDigit(current.Text[0]))
                    {
                        hfstring += " ";
                    }
                }
            }
            else if (current.FontSize != null && prev.FontSize == null)
            {
                hfstring += "&" + current.FontSize;
                if (current.Text != null && char.IsDigit(current.Text[0]))
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
                case ExcelHeaderFooterFormattingCodes.PageNumber:
                    hfstring += "&P" + current.PageNumberSuffix;
                    break;
                case ExcelHeaderFooterFormattingCodes.NumberOfPages:
                    hfstring += "&N";
                    break;
                case ExcelHeaderFooterFormattingCodes.SheetName:
                    hfstring += "&A";
                    break;
                case ExcelHeaderFooterFormattingCodes.FilePath:
                    hfstring += "&Z";
                    break;
                case ExcelHeaderFooterFormattingCodes.FileName:
                    hfstring += "&F";
                    break;
                case ExcelHeaderFooterFormattingCodes.CurrentDate:
                    hfstring += "&D";
                    break;
                case ExcelHeaderFooterFormattingCodes.CurrentTime:
                    hfstring += "&T";
                    break;
                case ExcelHeaderFooterFormattingCodes.Image:
                    hfstring += "&G";
                    if (current.Text == "&") hfstring += "&";
                    hfstring += current.Text;
                    break;
                case ExcelHeaderFooterFormattingCodes.Text:
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

        IEnumerator<ExcelHeaderFooterTextItem> IEnumerable<ExcelHeaderFooterTextItem>.GetEnumerator()
        {
            return _textCollection.GetEnumerator();
        }
    }
}
