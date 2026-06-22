/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlus.Export.Pdf.DocumentObjects.Fonts
{
    internal class PdfFontWidths : PdfObject
    {
        internal readonly int[] widths;

        public PdfFontWidths(int objectNumber, List<int> widths, int version = 0)
            : base(objectNumber, version)
        {
            this.widths = widths.ToArray();
        }

        //can Remove this construcotr probably.
        public PdfFontWidths(int objectNumber, Dictionary<FontMetricsClass, float> w, Dictionary<char, FontMetricsClass> m, int version = 0)
            : base(objectNumber, version)
        {
            float e = 2048;
            float d = 72f / 96f;
            float r = e * d;
            float k = 1000f;
            var l = (int)(w[m[' ']] * r * k / e);
            List<char> exsist = new List<char>();
            List<char> no = new List<char>();
            for (int i = 0; i <= 900; i++)
            {
                char c = (char)i;
                if (m.ContainsKey(c))
                {
                    exsist.Add(c);
                }
                else
                {
                    no.Add(c);
                }
            }
            //this.widths = new int[]
            //{
            //    (int)((w[m[' ']]*r*k)/e), (int)((w[m['!']]*r*k)/e), (int)((w[m['"']]*r*k)/e), (int)((w[m['#']]*r*k)/e), (int)((w[m['$']]*r*k)/e), (int)((w[m['%']]*r*k)/e), (int)((w[m['&']]*r*k)/e), (int)((w[m['\'']]*r*k)/e), (int)((w[m['(']]*r*k)/e), (int)((w[m[')']]*r*k)/e),
            //    (int)((w[m['*']]*r*k)/e), (int)((w[m['+']]*r*k)/e), (int)((w[m[',']]*r*k)/e), (int)((w[m['-']]*r*k)/e), (int)((w[m['.']]*r*k)/e), (int)((w[m[' ']]*r*k)/e), (int)((w[m['0']]*r*k)/e), (int)((w[m['1']]*r*k)/e), (int)((w[m['2']]*r*k)/e), (int)((w[m['3']]*r*k)/e),
            //    (int)((w[m['4']]*r*k)/e), (int)((w[m['5']]*r*k)/e), (int)((w[m['6']]*r*k)/e), (int)((w[m['7']]*r*k)/e), (int)((w[m['8']]*r*k)/e), (int)((w[m[' ']]*r*k)/e), (int)((w[m[':']]*r*k)/e), (int)((w[m[';']]*r*k)/e), (int)((w[m['<']]*r*k)/e), (int)((w[m['=']]*r*k)/e),
            //    (int)((w[m['>']]*r*k)/e), (int)((w[m['?']]*r*k)/e), (int)((w[m['@']]*r*k)/e), (int)((w[m['A']]*r*k)/e), (int)((w[m['B']]*r*k)/e), (int)((w[m['C']]*r*k)/e), (int)((w[m['D']]*r*k)/e), (int)((w[m['E']]*r*k)/e), (int)((w[m['F']]*r*k)/e), (int)((w[m['G']]*r*k)/e),
            //    (int)((w[m['H']]*r*k)/e), (int)((w[m['I']]*r*k)/e), (int)((w[m['J']]*r*k)/e), (int)((w[m['K']]*r*k)/e), (int)((w[m['L']]*r*k)/e), (int)((w[m['M']]*r*k)/e), (int)((w[m['N']]*r*k)/e), (int)((w[m['O']]*r*k)/e), (int)((w[m['P']]*r*k)/e), (int)((w[m['Q']]*r*k)/e),
            //    (int)((w[m['R']]*r*k)/e), (int)((w[m['S']]*r*k)/e), (int)((w[m[' ']]*r*k)/e), (int)((w[m['U']]*r*k)/e), (int)((w[m['V']]*r*k)/e), (int)((w[m['W']]*r*k)/e), (int)((w[m['X']]*r*k)/e), (int)((w[m['Y']]*r*k)/e), (int)((w[m['Z']]*r*k)/e), (int)((w[m['[']]*r*k)/e),
            //   (int)((w[m['\\']]*r*k)/e), (int)((w[m[']']]*r*k)/e), (int)((w[m['^']]*r*k)/e), (int)((w[m['_']]*r*k)/e), (int)((w[m['`']]*r*k)/e), (int)((w[m['a']]*r*k)/e), (int)((w[m['b']]*r*k)/e), (int)((w[m['c']]*r*k)/e), (int)((w[m['d']]*r*k)/e), (int)((w[m['e']]*r*k)/e),
            //    (int)((w[m['f']]*r*k)/e), (int)((w[m['g']]*r*k)/e), (int)((w[m['h']]*r*k)/e), (int)((w[m['i']]*r*k)/e), (int)((w[m['j']]*r*k)/e), (int)((w[m['k']]*r*k)/e), (int)((w[m['l']]*r*k)/e), (int)((w[m['m']]*r*k)/e), (int)((w[m['n']]*r*k)/e), (int)((w[m['o']]*r*k)/e),
            //    (int)((w[m['p']]*r*k)/e), (int)((w[m['q']]*r*k)/e), (int)((w[m['r']]*r*k)/e), (int)((w[m['s']]*r*k)/e), (int)((w[m['t']]*r*k)/e), (int)((w[m['u']]*r*k)/e), (int)((w[m[' ']]*r*k)/e), (int)((w[m['w']]*r*k)/e), (int)((w[m['x']]*r*k)/e), (int)((w[m[' ']]*r*k)/e),
            //    (int)((w[m['z']]*r*k)/e), (int)((w[m['{']]*r*k)/e), (int)((w[m['|']]*r*k)/e), (int)((w[m['}']]*r*k)/e), (int)((w[m[' ']]*r*k)/e), (int)((w[m[' ']]*r*k)/e), (int)((w[m['€']]*r*k)/e), (int)((w[m[' ']]*r*k)/e), (int)((w[m['‚']]*r*k)/e), (int)((w[m['ƒ']]*r*k)/e),
            //    (int)((w[m[' ']]*r*k)/e), (int)((w[m['„']]*r*k)/e), (int)((w[m['…']]*r*k)/e), (int)((w[m['†']]*r*k)/e), (int)((w[m['‡']]*r*k)/e), (int)((w[m['ˆ']]*r*k)/e), (int)((w[m['‰']]*r*k)/e), (int)((w[m['Š']]*r*k)/e), (int)((w[m['‹']]*r*k)/e), (int)((w[m['Œ']]*r*k)/e),
            //    (int)((w[m[' ']]*r*k)/e), (int)((w[m[' ']]*r*k)/e), (int)((w[m['Ž']]*r*k)/e), (int)((w[m[' ']]*r*k)/e), (int)((w[m[' ']]*r*k)/e), (int)((w[m['‘']]*r*k)/e), (int)((w[m['’']]*r*k)/e), (int)((w[m['“']]*r*k)/e), (int)((w[m['”']]*r*k)/e), (int)((w[m['•']]*r*k)/e),
            //    (int)((w[m['–']]*r*k)/e), (int)((w[m['—']]*r*k)/e), (int)((w[m['˜']]*r*k)/e), (int)((w[m['™']]*r*k)/e), (int)((w[m['š']]*r*k)/e), (int)((w[m['›']]*r*k)/e), (int)((w[m['œ']]*r*k)/e), (int)((w[m[' ']]*r*k)/e), (int)((w[m['ž']]*r*k)/e), (int)((w[m['Ÿ']]*r*k)/e),
            //    (int)((w[m['¡']]*r*k)/e), (int)((w[m['¢']]*r*k)/e), (int)((w[m['£']]*r*k)/e), (int)((w[m['¤']]*r*k)/e), (int)((w[m['¥']]*r*k)/e), (int)((w[m['¦']]*r*k)/e), (int)((w[m[' ']]*r*k)/e), (int)((w[m['¨']]*r*k)/e), (int)((w[m['©']]*r*k)/e), (int)((w[m['ª']]*r*k)/e),
            //    (int)((w[m['«']]*r*k)/e), (int)((w[m['¬']]*r*k)/e), (int)((w[m[' ']]*r*k)/e), (int)((w[m[' ']]*r*k)/e), (int)((w[m['¯']]*r*k)/e), (int)((w[m['°']]*r*k)/e), (int)((w[m['±']]*r*k)/e), (int)((w[m['²']]*r*k)/e), (int)((w[m['³']]*r*k)/e), (int)((w[m['´']]*r*k)/e),
            //    (int)((w[m['µ']]*r*k)/e), (int)((w[m['¶']]*r*k)/e), (int)((w[m['·']]*r*k)/e), (int)((w[m['¸']]*r*k)/e), (int)((w[m['¹']]*r*k)/e), (int)((w[m['º']]*r*k)/e), (int)((w[m['»']]*r*k)/e), (int)((w[m['¼']]*r*k)/e), (int)((w[m['½']]*r*k)/e), (int)((w[m['¾']]*r*k)/e),
            //    (int)((w[m['¿']]*r*k)/e), (int)((w[m['À']]*r*k)/e), (int)((w[m['Á']]*r*k)/e), (int)((w[m['Â']]*r*k)/e), (int)((w[m['Ã']]*r*k)/e), (int)((w[m['Ä']]*r*k)/e), (int)((w[m['Å']]*r*k)/e), (int)((w[m['Æ']]*r*k)/e), (int)((w[m['Ç']]*r*k)/e), (int)((w[m['È']]*r*k)/e),
            //    (int)((w[m['É']]*r*k)/e), (int)((w[m['Ê']]*r*k)/e), (int)((w[m['Ë']]*r*k)/e), (int)((w[m['Ì']]*r*k)/e), (int)((w[m['Í']]*r*k)/e), (int)((w[m['Î']]*r*k)/e), (int)((w[m['Ï']]*r*k)/e), (int)((w[m['Ð']]*r*k)/e), (int)((w[m['Ñ']]*r*k)/e), (int)((w[m['Ò']]*r*k)/e),
            //    (int)((w[m['Ó']]*r*k)/e), (int)((w[m['Ô']]*r*k)/e), (int)((w[m['Õ']]*r*k)/e), (int)((w[m['Ö']]*r*k)/e), (int)((w[m['×']]*r*k)/e), (int)((w[m['Ø']]*r*k)/e), (int)((w[m['Ù']]*r*k)/e), (int)((w[m['Ú']]*r*k)/e), (int)((w[m['Û']]*r*k)/e), (int)((w[m['Ü']]*r*k)/e),
            //    (int)((w[m['Ý']]*r*k)/e), (int)((w[m['Þ']]*r*k)/e), (int)((w[m['ß']]*r*k)/e), (int)((w[m['à']]*r*k)/e), (int)((w[m['á']]*r*k)/e), (int)((w[m['â']]*r*k)/e), (int)((w[m['ã']]*r*k)/e), (int)((w[m['ä']]*r*k)/e), (int)((w[m['å']]*r*k)/e), (int)((w[m['æ']]*r*k)/e),
            //    (int)((w[m['ç']]*r*k)/e), (int)((w[m['è']]*r*k)/e), (int)((w[m['é']]*r*k)/e), (int)((w[m['ê']]*r*k)/e), (int)((w[m['ë']]*r*k)/e), (int)((w[m['ì']]*r*k)/e), (int)((w[m['í']]*r*k)/e), (int)((w[m['î']]*r*k)/e), (int)((w[m['ï']]*r*k)/e), (int)((w[m['ð']]*r*k)/e),
            //    (int)((w[m['ñ']]*r*k)/e), (int)((w[m['ò']]*r*k)/e), (int)((w[m['ó']]*r*k)/e), (int)((w[m['ô']]*r*k)/e), (int)((w[m['õ']]*r*k)/e), (int)((w[m['ö']]*r*k)/e), (int)((w[m['÷']]*r*k)/e), (int)((w[m['ø']]*r*k)/e), (int)((w[m['ù']]*r*k)/e), (int)((w[m['ú']]*r*k)/e),
            //    (int)((w[m['û']]*r*k)/e), (int)((w[m['ü']]*r*k)/e), (int)((w[m[' ']]*r*k)/e), (int)((w[m['þ']]*r*k)/e), (int)((w[m['ÿ']]*r*k)/e)
            //};
            widths = new int[]
            {
                (int)(w[m[' ']]*r*k/e), (int)(w[m['!']]*r*k/e), (int)(w[m['"']]*r*k/e), (int)(w[m['#']]*r*k/e), (int)(w[m['$']]*r*k/e), (int)(w[m['%']]*r*k/e), (int)(w[m['&']]*r*k/e), (int)(w[m['\'']]*r*k/e), (int)(w[m['(']]*r*k/e), (int)(w[m[')']]*r*k/e),
                (int)(w[m['*']]*r*k/e), (int)(w[m['+']]*r*k/e), (int)(w[m[',']]*r*k/e), (int)(w[m['-']]*r*k/e), (int)(w[m['.']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m['0']]*r*k/e), (int)(w[m['1']]*r*k/e), (int)(w[m['2']]*r*k/e), (int)(w[m['3']]*r*k/e),
                (int)(w[m['4']]*r*k/e), (int)(w[m['5']]*r*k/e), (int)(w[m['6']]*r*k/e), (int)(w[m['7']]*r*k/e), (int)(w[m['8']]*r*k/e), (int)(w[m['8']]*r*k/e), (int)(w[m[':']]*r*k/e), (int)(w[m[';']]*r*k/e), (int)(w[m['<']]*r*k/e), (int)(w[m['=']]*r*k/e),
                (int)(w[m['>']]*r*k/e), (int)(w[m['?']]*r*k/e), (int)(w[m['@']]*r*k/e), (int)(w[m['A']]*r*k/e), (int)(w[m['B']]*r*k/e), (int)(w[m['C']]*r*k/e), (int)(w[m['D']]*r*k/e), (int)(w[m['E']]*r*k/e), (int)(w[m['F']]*r*k/e), (int)(w[m['G']]*r*k/e),
                (int)(w[m['H']]*r*k/e), (int)(w[m['I']]*r*k/e), (int)(w[m['J']]*r*k/e), (int)(w[m['K']]*r*k/e), (int)(w[m['L']]*r*k/e), (int)(w[m['M']]*r*k/e), (int)(w[m['N']]*r*k/e), (int)(w[m['O']]*r*k/e), (int)(w[m['P']]*r*k/e), (int)(w[m['Q']]*r*k/e),
                (int)(w[m['R']]*r*k/e), (int)(w[m['S']]*r*k/e), (int)(w[m['W']]*r*k/e), (int)(w[m['U']]*r*k/e), (int)(w[m['V']]*r*k/e), (int)(w[m['W']]*r*k/e), (int)(w[m['X']]*r*k/e), (int)(w[m['Y']]*r*k/e), (int)(w[m['Z']]*r*k/e), (int)(w[m['[']]*r*k/e),
               (int)(w[m['\\']]*r*k/e), (int)(w[m[']']]*r*k/e), (int)(w[m['^']]*r*k/e), (int)(w[m['_']]*r*k/e), (int)(w[m['`']]*r*k/e), (int)(w[m['a']]*r*k/e), (int)(w[m['b']]*r*k/e), (int)(w[m['c']]*r*k/e), (int)(w[m['d']]*r*k/e), (int)(w[m['e']]*r*k/e),
                (int)(w[m['f']]*r*k/e), (int)(w[m['g']]*r*k/e), (int)(w[m['h']]*r*k/e), (int)(w[m['i']]*r*k/e), (int)(w[m['j']]*r*k/e), (int)(w[m['k']]*r*k/e), (int)(w[m['l']]*r*k/e), (int)(w[m['m']]*r*k/e), (int)(w[m['n']]*r*k/e), (int)(w[m['o']]*r*k/e),
                (int)(w[m['p']]*r*k/e), (int)(w[m['q']]*r*k/e), (int)(w[m['r']]*r*k/e), (int)(w[m['s']]*r*k/e), (int)(w[m['t']]*r*k/e), (int)(w[m['u']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m['w']]*r*k/e), (int)(w[m['x']]*r*k/e), (int)(w[m[' ']]*r*k/e),
                (int)(w[m['z']]*r*k/e), (int)(w[m['{']]*r*k/e), (int)(w[m['|']]*r*k/e), (int)(w[m['}']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e),
                (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e),
                (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e),
                (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e),
                (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e),
                (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e),
                (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e),
                (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e),
                (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e),
                (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e),
                (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e),
                (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e),
                (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e),
                (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e), (int)(w[m[' ']]*r*k/e)
            };

        }

        internal override string RenderDictionary()
        {
            var widthsStr = string.Join(" ", widths.Select(w => w.ToString()).ToArray());
            return $"   [ {widthsStr} ]";
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            var widthsStr = string.Join(" ", widths.Select(w => w.ToString()).ToArray());
            WriteAscii(bw, $"   [ {widthsStr} ]");
        }
    }
}
