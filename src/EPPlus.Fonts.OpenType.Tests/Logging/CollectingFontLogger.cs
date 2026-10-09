using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tests
{
    /// <summary>
    /// Test logger that keeps every event it is asked to log. Thread-safe.
    /// </summary>
    public class CollectingFontLogger : IFontLogger
    {
        private readonly object _lock = new object();
        private readonly List<FontLogEvent> _events = new List<FontLogEvent>();
        private readonly FontLogSeverity _minimumSeverity;

        public CollectingFontLogger()
            : this(FontLogSeverity.Debug)
        {
        }

        public CollectingFontLogger(FontLogSeverity minimumSeverity)
        {
            _minimumSeverity = minimumSeverity;
        }

        public bool IsEnabled(FontLogSeverity severity)
        {
            return severity >= _minimumSeverity;
        }

        public void Log(FontLogEvent logEvent)
        {
            lock (_lock)
            {
                _events.Add(logEvent);
            }
        }

        /// <summary>A snapshot of all events logged so far, in order.</summary>
        public IList<FontLogEvent> Events
        {
            get
            {
                lock (_lock)
                {
                    return new List<FontLogEvent>(_events);
                }
            }
        }

        /// <summary>A snapshot of the events of one type, in order.</summary>
        public IList<FontLogEvent> GetEvents(FontLogEventType type)
        {
            var result = new List<FontLogEvent>();
            foreach (var e in Events)
            {
                if (e.Type == type)
                    result.Add(e);
            }
            return result;
        }
    }
}