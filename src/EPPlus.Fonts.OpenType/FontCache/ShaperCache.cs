/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/02/2026         EPPlus Software AB           Extracted from OpenTypeFontEngine
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.FontCache
{
    /// <summary>
    /// Per-thread cache of shapers for one engine. Storage only — it holds no policy and never
    /// decides which kind of shaper a request should get.
    ///
    /// The cache is per-thread because <see cref="TextShaper"/> is stateful per call: font
    /// tracking is reset at the start of every Shape, ShapeLight and ExtractCharWidths, and read
    /// back afterwards through GetUsedFonts. Sharing one instance across threads would interleave
    /// that state.
    ///
    /// Rendering and measurement shapers are kept in separate dictionaries. A measurement request
    /// may store a metrics-only shaper, and a shared dictionary keyed only on font name and
    /// subfamily would let a later rendering request find it.
    /// </summary>
    internal class ShaperCache
    {
        /// <summary>
        /// Keyed on the cache instance rather than the engine. Keying on the engine would leave
        /// this class depending on the one it was extracted from, and the engine already owns
        /// exactly one of these.
        /// </summary>
        [System.ThreadStatic]
        private static Dictionary<ShaperCache, Entries> _threadLocal;

        private class Entries
        {
            internal readonly Dictionary<string, TextShaper> Rendering =
                new Dictionary<string, TextShaper>();

            internal readonly Dictionary<string, ITextShaper> Measurement =
                new Dictionary<string, ITextShaper>();
        }

        internal bool TryGetRendering(string key, out TextShaper shaper)
        {
            return GetOrCreateEntries().Rendering.TryGetValue(key, out shaper);
        }

        internal void AddRendering(string key, TextShaper shaper)
        {
            GetOrCreateEntries().Rendering[key] = shaper;
        }

        internal bool TryGetMeasurement(string key, out ITextShaper shaper)
        {
            return GetOrCreateEntries().Measurement.TryGetValue(key, out shaper);
        }

        internal void AddMeasurement(string key, ITextShaper shaper)
        {
            GetOrCreateEntries().Measurement[key] = shaper;
        }

        /// <summary>
        /// Clears this cache for the calling thread only. Other threads see no entry for this
        /// cache on their next access and rebuild from scratch.
        /// </summary>
        internal void ClearCurrentThread()
        {
            if (_threadLocal != null)
                _threadLocal.Remove(this);
        }

        private Entries GetOrCreateEntries()
        {
            // [ThreadStatic] field initializers only run on the primary thread. Every other
            // thread sees null and must initialize on first use.
            if (_threadLocal == null)
                _threadLocal = new Dictionary<ShaperCache, Entries>();

            Entries entries;
            if (!_threadLocal.TryGetValue(this, out entries))
            {
                entries = new Entries();
                _threadLocal[this] = entries;
            }
            return entries;
        }
    }
}