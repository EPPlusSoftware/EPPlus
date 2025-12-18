using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Scanner
{
    /// <summary>
    /// Lightweight, immutable representation of a single font face.
    /// Used for caching and lookup during font scanning.
    /// Contains no open streams or heavy objects – only metadata and table directory.
    /// </summary>
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