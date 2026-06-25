/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
  05/06/2026         EPPlus Software AB           Added Clone() to support thread-safe FindBestMatch
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType.Scanner
{
    /// <summary>
    /// Lightweight representation of a single font face. Used for caching and lookup during
    /// font scanning. Contains no open streams or heavy objects – only metadata and table directory.
    /// </summary>
    /// <remarks>
    /// Instances cached by FontScannerCache are shared between callers. Any code that needs to
    /// set per-query state (such as IsExactMatch) must call Clone() first and mutate the copy —
    /// mutating the cached instance is not thread-safe.
    /// </remarks>
    public class FontFaceInfo
    {
        /// <summary>
        /// Full path to the font file on disk.
        /// </summary>
        public string FilePath { get; internal set; }

        /// <summary>
        /// The detected format of the containing file.
        /// </summary>
        public FontFormat Format { get; internal set; }

        /// <summary>
        /// Offset within the file where this font face starts.
        /// 0 for regular TTF/OTF files, greater than 0 for faces inside TTC collections.
        /// </summary>
        public long OffsetInFile { get; internal set; }

        /// <summary>
        /// File modification time used for cache invalidation.
        /// </summary>
        public DateTime LastWriteTimeUtc { get; internal set; }

        /// <summary>
        /// Font family name (e.g. "Arial", "Roboto").
        /// </summary>
        public string FamilyName { get; internal set; }

        /// <summary>
        /// Font subfamily name from name table (e.g. "Regular", "Bold Italic").
        /// </summary>
        public string SubfamilyName { get; internal set; }

        /// <summary>
        /// Full font Name from name table (e.g. "Roboto Regular")
        /// </summary>
        public string FullFontName { get; internal set; }

        /// <summary>
        /// Normalized subfamily as enum.
        /// </summary>
        public FontSubFamily Subfamily { get; internal set; }

        public ushort FsSelection { get; internal set; }

        /// <summary>
        /// True if this face was returned from a query that matched it exactly (by family name
        /// and subfamily). This is per-query state — never set on a cached instance directly;
        /// always Clone() first.
        /// </summary>
        public bool IsExactMatch { get; internal set; }

        /// <summary>
        /// True if this face is a variable font (i.e. the file contains an 'fvar' table).
        /// A variable font can only be relied upon to deliver its default named instance unless
        /// the variation tables are interpolated, which this library does not yet do. Matching
        /// therefore treats a variable face as capable of delivering only its default subfamily.
        /// </summary>
        public bool IsVariable { get; internal set; }

        /// <summary>
        /// Table directory for this face.
        /// </summary>
        public Dictionary<string, TableRecord> TableRecords { get; internal set; }

        /// <summary>
        /// Unique cache key: FilePath + "|" + OffsetInFile
        /// </summary>
        internal string CacheKey
        {
            get { return FilePath + "|" + OffsetInFile; }
        }

        internal FontFaceInfo()
        {
            TableRecords = new Dictionary<string, TableRecord>(StringComparer.Ordinal);
        }

        /// <summary>
        /// Creates a shallow copy of this instance. All value-type and string fields are copied;
        /// the TableRecords dictionary is shared by reference because it is populated once during
        /// scanning and never mutated afterwards.
        /// </summary>
        /// <remarks>
        /// Used by FontScannerV2.FindBestMatch to avoid mutating the cached instance with
        /// per-query state (IsExactMatch). Mutating a cached instance would cause a race
        /// condition between parallel callers.
        /// </remarks>
        internal FontFaceInfo Clone()
        {
            return new FontFaceInfo
            {
                FilePath = FilePath,
                Format = Format,
                OffsetInFile = OffsetInFile,
                LastWriteTimeUtc = LastWriteTimeUtc,
                FamilyName = FamilyName,
                SubfamilyName = SubfamilyName,
                FullFontName = FullFontName,
                Subfamily = Subfamily,
                FsSelection = FsSelection,
                IsExactMatch = IsExactMatch,
                IsVariable = IsVariable, // carry variable-font flag into per-query copy
                TableRecords = TableRecords, // shared by reference — never mutated post-scan
            };
        }

        public override string ToString()
        {
            return string.Format("{0} {1} → {2}{3}",
                FamilyName,
                SubfamilyName,
                Path.GetFileName(FilePath),
                OffsetInFile > 0 ? " [TTC]" : "");
        }
    }
}