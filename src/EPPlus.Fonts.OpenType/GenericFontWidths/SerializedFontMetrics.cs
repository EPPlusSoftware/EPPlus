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
  09/01/2026         EPPlus Software AB       Compact character mapping
 *************************************************************************************************/
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.GenericFontWidths
{
    /// <summary>
    /// Font metrics deserialized from a .fmtr file.
    ///
    /// The character to width class mapping used to be a Dictionary&lt;char, FontMetricsClass&gt;
    /// with one entry per character. The .fmtr files are only 222 kB uncompressed precisely
    /// because they store ranges, and expanding those ranges into dictionary entries on load
    /// threw that away: 138 617 entries across the library at roughly 16 bytes each came to
    /// 2.12 MB. Keeping the ranges as ranges brings that to about 335 kB.
    ///
    /// Lookup is a direct index for U+0020 to U+00FF, which covers nearly all real cell
    /// content, and a binary search over the sorted ranges and single characters beyond that.
    /// The Latin-1 table costs 224 bytes per font and keeps the common path at the same cost as
    /// the old hash lookup.
    ///
    /// Build with AddRange and AddCharacter, then call Seal before use.
    /// </summary>
    internal class SerializedFontMetrics
    {
        private const int LatinFirst = 0x20;
        private const int LatinLast = 0xFF;
        private const byte NotMapped = 0xFF;

        /// <summary>
        /// Number of width classes the format allows, and the size of the class width table.
        /// </summary>
        private const int MaxClasses = 32;

        private readonly float[] _classWidths = new float[MaxClasses];
        private readonly bool[] _classWidthSet = new bool[MaxClasses];

        // Populated while building, released by Seal.
        private List<ushort> _buildRangeStart = new List<ushort>();
        private List<ushort> _buildRangeEnd = new List<ushort>();
        private List<byte> _buildRangeClass = new List<byte>();
        private List<ushort> _buildSingleChar = new List<ushort>();
        private List<byte> _buildSingleClass = new List<byte>();

        // Sorted by start / character. Parallel arrays rather than an array of structs so that
        // the byte class does not get padded up to the alignment of the ushort.
        private ushort[] _rangeStart;
        private ushort[] _rangeEnd;
        private byte[] _rangeClass;
        private ushort[] _singleChar;
        private byte[] _singleClass;

        private byte[] _latin;

        public FontMetricsFamilies Family { get; set; }

        public FontSubFamilies SubFamily { get; set; }

        public ushort Version { get; set; }

        public uint FontKey { get; set; }

        /// <summary>
        /// Baseline to baseline distance, in em scaled by 96/72 like the class widths.
        ///
        /// For version 2 files this is derived from Ascender1em, Descender1em and LineGap1em
        /// rather than read from the file, so the three can never disagree with the total. The
        /// version 1 field is still read for version 1 files.
        /// </summary>
        public float LineHeight1em { get; set; }

        /// <summary>
        /// Distance from the baseline to the top of the line box, in the same unit as
        /// LineHeight1em.
        ///
        /// Version 1 files do not carry this, so it is approximated from the line height. The
        /// ascent share measured across the 101 shipped fonts runs from 0.7349 (Courier New) to
        /// 0.8289 (Tahoma) with a median of 0.8047, so any single constant is out by up to 7%
        /// of the font height - about 1.2px at 11pt. That spread is why version 2 exists.
        /// </summary>
        public float Ascender1em { get; set; }

        /// <summary>
        /// Distance from the baseline down to the bottom of the line box, in the same unit as
        /// LineHeight1em. Positive, unlike the OpenType convention, so that ascender plus
        /// descender plus line gap is the line height with no sign to get wrong.
        /// </summary>
        public float Descender1em { get; set; }

        /// <summary>
        /// Extra leading between lines, in the same unit as LineHeight1em. Zero for every font
        /// currently shipped - none of them sets USE_TYPO_METRICS, so the generator falls back
        /// to usWinAscent and usWinDescent, which span the full line box on their own. Carried
        /// anyway so a font that does set it needs no further format change.
        /// </summary>
        public float LineGap1em { get; set; }

        /// <summary>
        /// Ascent share used for version 1 files, which store only the total.
        /// </summary>
        internal const float Version1AscentRatio = 0.8047f;

        /// <summary>
        /// Fills Ascender1em, Descender1em and LineGap1em for a version 1 file by splitting the
        /// line height. Keeping the approximation here means consumers do not need to know
        /// which version the metrics came from.
        /// </summary>
        internal void ApproximateVerticalMetricsFromLineHeight()
        {
            Ascender1em = LineHeight1em * Version1AscentRatio;
            Descender1em = LineHeight1em - Ascender1em;
            LineGap1em = 0f;
        }

        public FontMetricsClass DefaultWidthClass { get; set; }

        /// <summary>
        /// Width of the default class, resolved once by Seal.
        /// </summary>
        public float DefaultWidth { get; private set; }

        /// <summary>
        /// True once Seal has run and the metrics are ready for lookup.
        /// </summary>
        public bool IsSealed
        {
            get { return _latin != null; }
        }

        #region Building

        internal void SetClassWidth(FontMetricsClass cls, float width)
        {
            var ix = (int)cls;
            if (ix < 0 || ix >= MaxClasses) return;
            _classWidths[ix] = width;
            _classWidthSet[ix] = true;
        }

        internal void AddRange(ushort start, ushort end, FontMetricsClass cls)
        {
            _buildRangeStart.Add(start);
            _buildRangeEnd.Add(end);
            _buildRangeClass.Add((byte)cls);
        }

        internal void AddCharacter(ushort character, FontMetricsClass cls)
        {
            _buildSingleChar.Add(character);
            _buildSingleClass.Add((byte)cls);
        }

        /// <summary>
        /// Sorts and freezes the mapping. Must be called before any lookup.
        /// </summary>
        internal void Seal()
        {
            if (IsSealed) return;

            DefaultWidth = GetClassWidth(DefaultWidthClass);

            SortByKey(_buildRangeStart, _buildRangeEnd, _buildRangeClass);
            _rangeStart = _buildRangeStart.ToArray();
            _rangeEnd = _buildRangeEnd.ToArray();
            _rangeClass = _buildRangeClass.ToArray();

            SortByKey(_buildSingleChar, null, _buildSingleClass);
            _singleChar = _buildSingleChar.ToArray();
            _singleClass = _buildSingleClass.ToArray();

            _buildRangeStart = null;
            _buildRangeEnd = null;
            _buildRangeClass = null;
            _buildSingleChar = null;
            _buildSingleClass = null;

            // Resolve the Latin-1 block once so the common path needs no search. Done after the
            // arrays are built so it goes through the same lookup and cannot disagree with it.
            _latin = new byte[LatinLast - LatinFirst + 1];
            for (var c = LatinFirst; c <= LatinLast; c++)
            {
                byte cls;
                _latin[c - LatinFirst] = TrySearch((ushort)c, out cls) ? cls : NotMapped;
            }
        }

        /// <summary>
        /// Insertion sort over the parallel lists, keyed on the first. The lists arrive close to
        /// sorted from the file, and the alternative - building index arrays and permuting -
        /// allocates more than the sort saves at these sizes.
        /// </summary>
        private static void SortByKey(List<ushort> keys, List<ushort> second, List<byte> classes)
        {
            for (var i = 1; i < keys.Count; i++)
            {
                var key = keys[i];
                var sec = second == null ? (ushort)0 : second[i];
                var cls = classes[i];
                var j = i - 1;
                while (j >= 0 && keys[j] > key)
                {
                    keys[j + 1] = keys[j];
                    if (second != null) second[j + 1] = second[j];
                    classes[j + 1] = classes[j];
                    j--;
                }
                keys[j + 1] = key;
                if (second != null) second[j + 1] = sec;
                classes[j + 1] = cls;
            }
        }

        #endregion

        #region Lookup

        /// <summary>
        /// Width of a class, or the default width when the class has none.
        /// </summary>
        public float GetClassWidth(FontMetricsClass cls)
        {
            var ix = (int)cls;
            if (ix >= 0 && ix < MaxClasses && _classWidthSet[ix])
            {
                return _classWidths[ix];
            }
            return 0f;
        }

        /// <summary>
        /// True when the character has an explicit width class in this font.
        /// </summary>
        public bool ContainsCharacter(char c)
        {
            FontMetricsClass cls;
            return TryGetClass(c, out cls);
        }

        /// <summary>
        /// Resolves the width class of a character.
        /// </summary>
        public bool TryGetClass(char c, out FontMetricsClass cls)
        {
            if (c >= LatinFirst && c <= LatinLast)
            {
                var mapped = _latin[c - LatinFirst];
                if (mapped == NotMapped)
                {
                    cls = DefaultWidthClass;
                    return false;
                }
                cls = (FontMetricsClass)mapped;
                return true;
            }

            byte found;
            if (TrySearch(c, out found))
            {
                cls = (FontMetricsClass)found;
                return true;
            }
            cls = DefaultWidthClass;
            return false;
        }

        /// <summary>
        /// Width of a character, falling back to the default class width. This is the whole hot
        /// path in one call - it replaces a lookup in CharMetrics followed by a second one in
        /// ClassWidths.
        /// </summary>
        public float GetCharacterWidth(char c)
        {
            FontMetricsClass cls;
            if (TryGetClass(c, out cls))
            {
                return GetClassWidth(cls);
            }
            return DefaultWidth;
        }

        private bool TrySearch(ushort c, out byte cls)
        {
            // Singles first: they outnumber the ranges roughly twenty to one.
            if (_singleChar.Length > 0)
            {
                var lo = 0;
                var hi = _singleChar.Length - 1;
                while (lo <= hi)
                {
                    var mid = lo + ((hi - lo) >> 1);
                    var v = _singleChar[mid];
                    if (v == c)
                    {
                        cls = _singleClass[mid];
                        return true;
                    }
                    if (v < c) lo = mid + 1; else hi = mid - 1;
                }
            }

            if (_rangeStart.Length > 0)
            {
                var lo = 0;
                var hi = _rangeStart.Length - 1;
                while (lo <= hi)
                {
                    var mid = lo + ((hi - lo) >> 1);
                    if (_rangeStart[mid] > c)
                    {
                        hi = mid - 1;
                    }
                    else if (_rangeEnd[mid] < c)
                    {
                        lo = mid + 1;
                    }
                    else
                    {
                        cls = _rangeClass[mid];
                        return true;
                    }
                }
            }

            cls = 0;
            return false;
        }

        #endregion

        public uint GetKey()
        {
            return GetKey(Family, SubFamily);
        }

        public static uint GetKey(FontMetricsFamilies family, FontSubFamilies subFamily)
        {
            var k1 = (ushort)family;
            var k2 = (ushort)subFamily;
            return (uint)((k1 << 16) | ((k2) & 0xffff));
        }
    }
}