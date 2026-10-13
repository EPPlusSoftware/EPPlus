/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/08/2026         EPPlus Software AB           Font logging
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.IO;
using System.Text;

namespace EPPlus.Fonts.OpenType.Logging
{
    /// <summary>
    /// Writes <see cref="FontLogEvent.Message"/> as one line per event to a text file.
    /// The file is opened for each write and closed again, so no handle is held between events
    /// and the log can be read, moved or deleted while rendering is in progress. This logger
    /// therefore needs no disposal.
    /// </summary>
    internal sealed class TextFileFontLogger : IFontLogger
    {
        private readonly string _path;
        private readonly FontLogSeverity _minimumSeverity;
        private readonly object _lock = new object();

        internal TextFileFontLogger(FileInfo logfile, FontLogSeverity minimumSeverity)
        {
            if (logfile == null)
                throw new ArgumentNullException("logfile");

            _path = logfile.FullName;
            _minimumSeverity = minimumSeverity;

            // Start a fresh log, and fail here rather than during rendering if the path is unusable.
            using (new FileStream(_path, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
            {
            }
        }

        public bool IsEnabled(FontLogSeverity severity)
        {
            return severity >= _minimumSeverity;
        }

        public void Log(FontLogEvent logEvent)
        {
            if (logEvent == null || !IsEnabled(logEvent.Severity))
                return;

            var line = Prefix(logEvent.Severity) + " " + logEvent.Message;

            lock (_lock)
            {
                using (var stream = new FileStream(_path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                {
                    writer.WriteLine(line);
                }
            }
        }

        private static string Prefix(FontLogSeverity severity)
        {
            switch (severity)
            {
                case FontLogSeverity.Debug: return "DBG";
                case FontLogSeverity.Information: return "INF";
                case FontLogSeverity.Warning: return "WRN";
                default: return "ERR";
            }
        }
    }
}