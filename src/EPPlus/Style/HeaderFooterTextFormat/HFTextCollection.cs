using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.Packaging.Ionic.Zip;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography.Pkcs;
using System.Text;
namespace OfficeOpenXml.Style.HeaderFooterTextFormat
{
    public class HFTextCollection : IEnumerable<HFText>
    {
        internal ExcelWorkbook _wb;
        internal enum Lanes
        {
            Left,
            Center,
            Right,
        }

        internal Lanes lane;


        private List<HFText> _textCollection = new List<HFText>();

        //public List<HFText> TextCollection
        //{
        //    get { return _textCollection; }
        //}

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

        internal HFTextCollection(ExcelWorkbook wb, Lanes lane)
        {
            _wb = wb;
            this.lane = lane;
        }

        public void Add(HFText textItem)
        {
            _textCollection.Add(textItem);
        }
        public HFText Add(string text)
        {
            HFText hfText = new HFText(text);
            _textCollection.Add(hfText);
            return hfText;
        }
        public HFText Insert(int index, string text)
        {
            if(index < 0 || index >= _textCollection.Count) throw new IndexOutOfRangeException("index was out of range.");
            if (text == null) throw new ArgumentException("Text can't be null", "text");
            HFText hFText = new HFText();
            hFText.Text = text;
            _textCollection.Insert(index, hFText);
            return hFText;
        }
        public void Remove(HFText Item)
        {
            _textCollection.Remove(Item);
        }

        public void RemoveAt(int index)
        {
            _textCollection.RemoveAt(index);
        }
        public void Clear()
        {
            _textCollection.Clear();
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
                    Add(value);
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
        //IEnumerable

        private Color GetColor(string c, ref eThemeSchemeColor? theme, ref double? tint)
        {
            if (c.Contains('-') || c.Contains('+'))
            {
                theme = (eThemeSchemeColor)int.Parse(""+c[0] + c[1]);
                var themeColor = _wb.ThemeManager.GetOrCreateTheme().ColorScheme.GetColorByEnum((eThemeSchemeColor)theme);
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
            int i = 0;
            while (i < hfText.Length)
            {
                if (hfText[i] == '&')
                {
                    i++;
                    switch (hfText[i])
                    {
                        case 'L':
                            if (lane != Lanes.Left) stop = true;
                            break;
                        case 'C':
                            if (lane != Lanes.Center) stop = true;
                            break;
                        case 'R':
                            if (lane != Lanes.Right) stop = true;
                            break;
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
                            temp.theme = theme;
                            temp.tint = tint;
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
                    HFText newText = new HFText
                    {
                        Text = temp.Text,
                        Bold = temp.Bold,
                        Italic = temp.Italic,
                        Underline = temp.Underline,
                        DoubleUnderline = temp.DoubleUnderline,
                        Shadow = temp.Shadow,
                        Color = temp.Color,
                        theme = temp.theme,
                        tint = temp.tint,
                        Outline = temp.Outline,
                        Striketrough = temp.Striketrough,
                        SuperScript = temp.SuperScript,
                        SubScript = temp.SubScript,
                        FontSize = temp.FontSize,
                        FontName = temp.FontName,
                    };
                    _textCollection.Add(newText);
                    if (quoteTag)
                    {
                        temp = new HFText();
                        quoteTag = false;
                    }
                }
                if (stop) break;
            }
        }

        internal string WriteHeaderFooterFormat()
        {
            string hfstring = "";
            HFText prev = null;
            foreach (HFText text in _textCollection)
            {
                hfstring += WriteHeaderFooter2(text, prev);
            }
            return hfstring;
        }

        private string WriteHeaderFooter2(HFText current, HFText prev)
        {
            string hfstring = "";
            if (!string.IsNullOrEmpty(current.FontName) || current.FontName == "-" || current.Bold || current.Italic)
            {
                var fontParts = new List<string> { current.FontName ?? "-" };

                if (current.Bold) fontParts.Add("Bold");
                if (current.Italic) fontParts.Add("Italic");

                if (!current.Bold && !current.Italic)
                    fontParts.Add("Regular");

                hfstring += $"&\"{string.Join(",", fontParts.ToArray())}\"";
            }

            if (current.Underline) hfstring += "&U";
            if (current.DoubleUnderline) hfstring += "&E";
            if (current.Shadow) hfstring += "&H";

            if (current.theme != null) hfstring += $"&K{((int)current.theme).ToString("00")}{(current.tint >= 0 ? "+" : "-")}{Math.Abs((double)current.tint*100).ToString("000")}";
            else if (current.Color != Color.Empty) hfstring += $"&K{current.Color.R:X2}{current.Color.G:X2}{current.Color.B:X2}";

            if (current.Outline) hfstring += "&O";
            if (current.Striketrough) hfstring += "&S";
            if (current.SuperScript) hfstring += "&X";
            if (current.SubScript) hfstring += "&Y";
            if (current.FontSize != null) hfstring += "&" + current.FontSize + " ";
            hfstring += current.Text;
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
