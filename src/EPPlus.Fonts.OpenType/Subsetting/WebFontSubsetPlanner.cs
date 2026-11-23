/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/19/2026         EPPlus Software AB           Web font subsetting for SVG export
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Logging;
using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.Utils;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    /// <summary>
    /// Plans the fonts embedded in web output (SVG/HTML): which subsets to declare with @font-face,
    /// and the font-family stack for each font the text was measured with.
    ///
    /// For every added font R (the measured family) the planner finds the font M that measurement
    /// used, routes each code point with <see cref="DefaultFontProvider"/> exactly as measurement
    /// does, decides per font whether it may be embedded (web policy and OnFontEmbedding callback),
    /// and subsets every embedded font once with the union of the code points routed to it.
    ///
    /// The stack for R:
    ///   M may be embedded:      subset(M), glyph fallbacks, R, R's fallback chain
    ///   M may not be embedded:  M, subset(S), glyph fallbacks, R, R's fallback chain
    /// followed by a generic family (<see cref="WebFontPlan.GetGenericFamily"/>). S is a safety net
    /// carrying the code points routed to M: the first font of R's chain that is installed and may be
    /// embedded, otherwise the bundled Archivo Narrow. Archivo Narrow is never written by name, since
    /// browsers do not have it; it is only used embedded.
    ///
    /// A font is not embedded when the policy skips it, when it has CFF outlines, when the policy or
    /// callback asks for the whole font (whole fonts are never embedded in web output), or when
    /// subsetting it fails. When measurement used serialized metrics because R is not installed,
    /// M is R itself and is written by name.
    ///
    /// Glyph fallbacks are the fonts the provider routed code points to, in priority order (Noto
    /// Emoji, Noto Math, then script fonts). A fallback that may not be embedded is written by name,
    /// without a safety net, followed by the other fonts of its script's fallback chain so a viewer
    /// without it still gets a font for the script. Code points that no font has are left out of
    /// every subset; the browser picks a system font for them.
    ///
    /// A browser walks the stack per character and takes the first font that has it, which for fonts
    /// written by name need not be the font measurement used: a Chinese font early in the stack may
    /// also have the kana that were measured with a Japanese font. Code points routed to a fallback
    /// written by name therefore get a stack of their own that starts with that fallback, see
    /// <see cref="WebFontPlan.GetFamilyNames(string, FontSubFamily, int)"/>. Embedded subsets need
    /// no such stack: a subset only has the code points routed to it.
    /// 
    /// A planner created with embedFonts false embeds nothing: every font is written by name, there is
    /// no safety net, and the stack starts with R rather than M. Stacks and the stacks per code point
    /// are planned the same way, so text can still be split where measurement used a glyph fallback.
    ///
    ///
    /// One instance plans one output. Not thread-safe.
    /// </summary>
    public sealed class WebFontSubsetPlanner
    {
        private const string LastResortFamily = "Archivo Narrow";
        private readonly OpenTypeFontEngine _engine;
        private readonly bool _embedFonts;
        private readonly List<Request> _requests = new List<Request>();
        private readonly Dictionary<string, Request> _requestByKey = new Dictionary<string, Request>(StringComparer.Ordinal);

        // Whether a font may be embedded, cached per font so the OnFontEmbedding callback runs once per font.
        private readonly Dictionary<FontKey, bool> _embeddable = new Dictionary<FontKey, bool>();
        private readonly HashSet<FontKey> _subsetFailed = new HashSet<FontKey>();
        private bool _built;

        /// <summary>
        /// Creates a planner that embeds the fonts the web policy and the OnFontEmbedding callback allow.
        /// </summary>
        public WebFontSubsetPlanner(OpenTypeFontEngine engine)
            : this(engine, true)
        {
        }

        /// <summary>
        /// Creates a planner.
        /// </summary>
        /// <param name="engine">The engine the text was measured with.</param>
        /// <param name="embedFonts">
        /// False to embed nothing and write every font by name. The embedding policy and the
        /// OnFontEmbedding callback are then never consulted.
        /// </param>
        public WebFontSubsetPlanner(OpenTypeFontEngine engine, bool embedFonts)
        {
            if (engine == null)
                throw new ArgumentNullException("engine");
            _engine = engine;
            _embedFonts = embedFonts;
        }

        /// <summary>
        /// Adds text measured with the given family and subfamily. Call once per text run; text for the
        /// same font is collected. An empty text only registers the font, so it still gets a stack
        /// (for instance a paragraph's default font, which has no text of its own).
        /// </summary>
        /// <param name="family">The family the text was measured with (after any web font substitution).</param>
        /// <param name="subFamily">The subfamily the text was measured with.</param>
        /// <param name="text">The text, or null or empty to only register the font.</param>
        public void AddText(string family, FontSubFamily subFamily, string text)
        {
            if (_built)
                throw new InvalidOperationException("Cannot add text after Build().");
            if (string.IsNullOrEmpty(family))
                throw new ArgumentException("A family name is required.", "family");

            var key = RequestKey(family, subFamily);
            Request request;
            if (!_requestByKey.TryGetValue(key, out request))
            {
                request = new Request(family, subFamily);
                _requestByKey.Add(key, request);
                _requests.Add(request);
            }

            if (string.IsNullOrEmpty(text))
                return;

            foreach (var cp in CodePointUtil.ExtractCodePoints(text))
                request.CodePoints.Add(cp);
        }

        /// <summary>
        /// Routes, decides and subsets. Can be called once.
        /// </summary>
        /// <summary>
        /// Routes, decides and subsets. Can be called once.
        /// </summary>
        public WebFontPlan Build()
        {
            if (_built)
                throw new InvalidOperationException("Build() can only be called once.");
            _built = true;

            foreach (var request in _requests)
                Route(request);

            // A failed subset marks its font as not embeddable and reruns the assignment, so a failed
            // primary is written by name followed by a safety net (ultimately the bundled Archivo
            // Narrow), and a failed glyph fallback is written by name. The loop ends: a font that
            // failed is never assigned again, and there are finitely many fonts. When it ends, every
            // font the assignment embeds has a subset.
            while (true)
            {
                var assignment = Assign();
                var failed = new List<FontKey>();
                var faces = SubsetAll(assignment, failed);

                if (failed.Count == 0)
                    return Compose(assignment, faces);

                foreach (var id in failed)
                    _subsetFailed.Add(id);
            }
        }

        internal static string RequestKey(string family, FontSubFamily subFamily)
        {
            return (family ?? string.Empty).ToUpperInvariant() + "|" + (int)subFamily;
        }

        // -----------------------------------------------------------------------------------------
        // Routing: which font each code point is drawn with, as measurement routes it
        // -----------------------------------------------------------------------------------------

        private void Route(Request request)
        {
            // Measurement uses serialized metrics, not a font file, when R is not installed and
            // resolution fell through to the last-resort font. M is then R itself.
            var shaper = _engine.GetMeasurementShaper(request.Family, request.SubFamily);
            request.MeasuredFromMetrics = shaper == null || !shaper.HasGlyphIds;

            // Routing needs a real font even when measurement used metrics: the code points that
            // stay with the primary are the ones M (or its safety net) must carry.
            request.Primary = _engine.LoadFont(request.Family, request.SubFamily);
            if (request.Primary == null)
            {
                request.PrimaryCodePoints.UnionWith(request.CodePoints);
                return;
            }

            var provider = new DefaultFontProvider(_engine, request.Primary);
            var sorted = new List<int>(request.CodePoints);
            sorted.Sort();

            foreach (var cp in sorted)
            {
                OpenTypeFont dest;
                ushort glyphId;
                if (!provider.TryGetGlyphFont((uint)cp, out dest, out glyphId))
                    continue; // no font has it

                if (ReferenceEquals(dest, request.Primary))
                {
                    request.PrimaryCodePoints.Add(cp);
                    continue;
                }

                var id = IdentityOf(dest);
                Fallback fallback;
                if (!request.FallbackById.TryGetValue(id, out fallback))
                {
                    fallback = new Fallback(dest, id, Rank(cp), cp);
                    request.FallbackById.Add(id, fallback);
                    request.Fallbacks.Add(fallback);
                }
                fallback.CodePoints.Add(cp);
                var script = UnicodeScriptClassifier.OfCodePoint((uint)cp);
                if (!fallback.Scripts.Contains(script))
                    fallback.Scripts.Add(script);
            }

            // Priority order of the provider: emoji, math, then script fonts.
            request.Fallbacks.Sort((a, b) =>
                a.Rank != b.Rank ? a.Rank.CompareTo(b.Rank) : a.FirstCodePoint.CompareTo(b.FirstCodePoint));
        }

        private static int Rank(int codePoint)
        {
            switch (UnicodeScriptClassifier.OfCodePoint((uint)codePoint))
            {
                case UnicodeScript.Emoji:
                    return 0;
                case UnicodeScript.Math:
                    return 1;
                default:
                    return 2;
            }
        }

        // -----------------------------------------------------------------------------------------
        // Assignment: which fonts are embedded and which code points each subset carries
        // -----------------------------------------------------------------------------------------

        private Assignment Assign()
        {
            var assignment = new Assignment();

            foreach (var request in _requests)
            {
                var slot = new Slot(request);

                if (_embedFonts && !request.MeasuredFromMetrics && request.Primary != null && CanEmbed(request.Primary))
                {
                    slot.PrimaryFace = IdentityOf(request.Primary);
                    assignment.Add(request.Primary, request.PrimaryCodePoints);

                    if (IsLastResort(request.Primary.GetEnglishFontFamilyName()) && !IsLastResort(request.Family))
                    {
                        LogWarning(request.Family, LastResortFamily, string.Format(
                            "'{0}' {1} and its fallback chain are not installed; the text was measured with the last-resort font '{2}', which is embedded in the web output.",
                            request.Family, request.SubFamily, LastResortFamily));
                    }
                }
                else
                {
                    // Without embedding the requested family comes first, as the fonts are made
                    // available by other means; the font measurement resolved to follows in its chain.
                    slot.PrimaryName = !_embedFonts || request.MeasuredFromMetrics || request.Primary == null
                        ? request.Family
                        : request.Primary.GetEnglishFontFamilyName();

                    var safetyNet = _embedFonts ? FindSafetyNet(request) : null;
                    if (safetyNet != null)
                    {
                        slot.SafetyNet = safetyNet;
                        assignment.Add(safetyNet, request.PrimaryCodePoints);
                    }
                }

                foreach (var fallback in request.Fallbacks)
                {
                    if (_embedFonts && CanEmbed(fallback.Font))
                    {
                        slot.Fallbacks.Add(new StackEntry(fallback, fallback.Id, null));
                        assignment.Add(fallback.Font, fallback.CodePoints);
                    }
                    else
                    {
                        slot.Fallbacks.Add(new StackEntry(fallback, null, fallback.Font.GetEnglishFontFamilyName()));
                    }
                }

                assignment.Slots.Add(slot);
            }

            return assignment;
        }

        /// <summary>
        /// The first font of R's fallback chain that is installed and may be embedded, otherwise the
        /// bundled Archivo Narrow. Null only if Archivo Narrow cannot be embedded either.
        /// </summary>
        private OpenTypeFont FindSafetyNet(Request request)
        {
            foreach (var name in _engine.GetFallbackChain(request.Family))
            {
                if (_engine.GetFontAvailability(name, request.SubFamily) != FontAvailability.Exact)
                    continue;

                var candidate = _engine.LoadFont(name, request.SubFamily);
                if (candidate != null && CanEmbed(candidate))
                    return candidate;
            }

            var lastResort = EmbeddedFonts.LoadArchivoNarrow(request.SubFamily);
            if (CanEmbed(lastResort))
            {
                LogWarning(request.Family, LastResortFamily, string.Format(
                    "'{0}' {1} may not be embedded and no font in its fallback chain is installed and embeddable; '{2}' is embedded as the safety net.",
                    request.Family, request.SubFamily, LastResortFamily));
                return lastResort;
            }
            return null;
        }

        private bool CanEmbed(OpenTypeFont font)
        {
            var id = IdentityOf(font);
            if (_subsetFailed.Contains(id))
                return false;

            bool embeddable;
            if (_embeddable.TryGetValue(id, out embeddable))
                return embeddable;

            if (font.Format != FontFormat.Ttf)
            {
                LogInformation(id.Family, string.Format(
                    "'{0}' {1} has CFF outlines, which are not subsetted for web output; it is referenced by name.",
                    id.Family, id.SubFamily));
                embeddable = false;
            }
            else
            {
                var decision = _engine.ResolveEmbeddingDecision(font, FontRenderTarget.Web);
                if (decision == FontEmbeddingDecision.EmbedWhole)
                {
                    LogInformation(id.Family, string.Format(
                        "'{0}' {1}: EmbedWhole was requested, but whole fonts are never embedded in web output; it is referenced by name. Return Subset from OnFontEmbedding to embed it.",
                        id.Family, id.SubFamily));
                }
                embeddable = decision == FontEmbeddingDecision.Subset;
            }

            _embeddable[id] = embeddable;
            return embeddable;
        }

        // -----------------------------------------------------------------------------------------
        // Subsetting
        // -----------------------------------------------------------------------------------------

        private Dictionary<FontKey, WebFontFace> SubsetAll(Assignment assignment, List<FontKey> failed)
        {
            var faces = new Dictionary<FontKey, WebFontFace>();

            foreach (var id in assignment.Order)
            {
                var font = assignment.FontById[id];
                var codePoints = assignment.CodePointsById[id];

                try
                {
                    // Web subsets are unhinted: browsers on most platforms ignore TrueType hinting, text is
                    // laid out with geometric precision and a fixed textLength, and the hinting programs
                    // are often the largest part of a small subset.
                    var subset = font.CreateSubset(CodePointUtil.CodePointsToString(codePoints),
                        new SubsetOptions { AddSubsetTag = true, RemoveHinting = true });
                    var data = subset.Serialize();
                    var familyName = SubsetTag.Create(id, codePoints) + id.Family;
                    faces.Add(id, new WebFontFace(familyName, id.Family, id.SubFamily, data));
                }
                catch (Exception ex)
                {
                    LogWarning(id.Family, null, string.Format(
                        "Subsetting '{0}' {1} for web output failed ({2}: {3}); it is referenced by name instead.",
                        id.Family, id.SubFamily, ex.GetType().Name, ex.Message));
                    failed.Add(id);
                }
            }

            return faces;
        }

        // -----------------------------------------------------------------------------------------
        // Stacks
        // -----------------------------------------------------------------------------------------

        private WebFontPlan Compose(Assignment assignment, Dictionary<FontKey, WebFontFace> faces)
        {
            var stacks = new Dictionary<string, WebFontPlan.Stack>(StringComparer.Ordinal);

            foreach (var slot in assignment.Slots)
            {
                var request = slot.Request;
                var names = new List<string>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                if (slot.PrimaryFace.HasValue)
                {
                    AddFace(names, seen, faces, slot.PrimaryFace.Value);
                }
                else
                {
                    AddName(names, seen, slot.PrimaryName);
                    if (slot.SafetyNet != null)
                        AddFace(names, seen, faces, IdentityOf(slot.SafetyNet));
                }

                var byName = new List<StackEntry>();
                foreach (var entry in slot.Fallbacks)
                {
                    if (entry.Face.HasValue)
                    {
                        AddFace(names, seen, faces, entry.Face.Value);
                    }
                    else
                    {
                        AddName(names, seen, entry.Name);
                        AddScriptChains(names, seen, entry.Source);
                        byName.Add(entry);
                    }
                }

                AddName(names, seen, request.Family);
                foreach (var name in _engine.GetFallbackChain(request.Family))
                    AddName(names, seen, name);

                stacks[RequestKey(request.Family, request.SubFamily)] =
                    new WebFontPlan.Stack(names, GetGenericFamily(slot), ComposeCodePointStacks(byName, names));
            }

            var orderedFaces = new List<WebFontFace>();
            foreach (var id in assignment.Order)
            {
                WebFontFace face;
                if (faces.TryGetValue(id, out face))
                    orderedFaces.Add(face);
            }

            return new WebFontPlan(orderedFaces, stacks);
        }

        /// <summary>
        /// Stacks for code points routed to fallbacks written by name: the fallback, the rest of its
        /// scripts' fallback chains, then the font's own stack. Null when there are none.
        /// </summary>
        private Dictionary<int, IReadOnlyList<string>> ComposeCodePointStacks(List<StackEntry> byName, List<string> fontStack)
        {
            if (byName.Count == 0)
                return null;

            var result = new Dictionary<int, IReadOnlyList<string>>();
            foreach (var entry in byName)
            {
                var names = new List<string>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                AddName(names, seen, entry.Name);
                AddScriptChains(names, seen, entry.Source);
                foreach (var name in fontStack)
                    AddName(names, seen, name);

                var stack = names.AsReadOnly();
                foreach (var cp in entry.Source.CodePoints)
                    result[cp] = stack;
            }
            return result;
        }

        /// <summary>
        /// Adds the configured fallback chain of every script the fallback was used for.
        /// </summary>
        private void AddScriptChains(List<string> names, HashSet<string> seen, Fallback fallback)
        {
            foreach (var script in fallback.Scripts)
            {
                var chain = _engine.FontStore.GetScriptFallback(script);
                if (chain == null)
                    continue;
                foreach (var name in chain)
                    AddName(names, seen, name);
            }
        }

        /// <summary>
        /// The generic family of the font M stands for. When M is R itself (measured from metrics),
        /// the safety net from R's chain represents it, being a metric-compatible substitute, or,
        /// without a safety net, the font R resolves to.
        /// </summary>
        private static GenericFontFamily GetGenericFamily(Slot slot)
        {
            var request = slot.Request;
            if (!request.MeasuredFromMetrics && request.Primary != null)
                return GenericFontFamilyClassifier.Classify(request.Primary);

            if (slot.SafetyNet != null && !IsLastResort(slot.SafetyNet.GetEnglishFontFamilyName()))
                return GenericFontFamilyClassifier.Classify(slot.SafetyNet);

            // Without a safety net (fonts are not embedded), the font R resolves to stands for it.
            if (request.Primary != null && !IsLastResort(request.Primary.GetEnglishFontFamilyName()))
                return GenericFontFamilyClassifier.Classify(request.Primary);

            return GenericFontFamily.SansSerif;
        }

        private static void AddFace(List<string> names, HashSet<string> seen, Dictionary<FontKey, WebFontFace> faces, FontKey id)
        {
            WebFontFace face;
            if (faces.TryGetValue(id, out face) && seen.Add(face.FamilyName))
                names.Add(face.FamilyName);
        }

        private static void AddName(List<string> names, HashSet<string> seen, string name)
        {
            // Archivo Narrow is only ever used embedded: browsers do not have it installed.
            if (string.IsNullOrEmpty(name) || IsLastResort(name))
                return;
            if (seen.Add(name))
                names.Add(name);
        }

        // -----------------------------------------------------------------------------------------
        // Helpers
        // -----------------------------------------------------------------------------------------

        private static FontKey IdentityOf(OpenTypeFont font)
        {
            return new FontKey(font.GetEnglishFontFamilyName(), font.NameTable.GetSubfamilyEnum());
        }

        private static bool IsLastResort(string family)
        {
            return string.Equals(LastResortFamily, family, StringComparison.OrdinalIgnoreCase);
        }

        private void LogWarning(string requested, string resolved, string message)
        {
            var logger = _engine.Logger;
            if (FontLog.IsEnabled(logger, FontLogSeverity.Warning))
                FontLog.Write(logger, FontLogSeverity.Warning, FontLogEventType.EmbeddingDecision, message, requested, resolved);
        }

        private void LogInformation(string requested, string message)
        {
            var logger = _engine.Logger;
            if (FontLog.IsEnabled(logger, FontLogSeverity.Information))
                FontLog.Write(logger, FontLogSeverity.Information, FontLogEventType.EmbeddingDecision, message, requested, null);
        }

        // -----------------------------------------------------------------------------------------
        // Working data
        // -----------------------------------------------------------------------------------------

        private sealed class Request
        {
            internal Request(string family, FontSubFamily subFamily)
            {
                Family = family;
                SubFamily = subFamily;
            }

            internal string Family { get; private set; }
            internal FontSubFamily SubFamily { get; private set; }
            internal HashSet<int> CodePoints { get; } = new HashSet<int>();

            internal bool MeasuredFromMetrics { get; set; }
            internal OpenTypeFont Primary { get; set; }
            internal HashSet<int> PrimaryCodePoints { get; } = new HashSet<int>();
            internal List<Fallback> Fallbacks { get; } = new List<Fallback>();
            internal Dictionary<FontKey, Fallback> FallbackById { get; } = new Dictionary<FontKey, Fallback>();
        }

        private sealed class Fallback
        {
            internal Fallback(OpenTypeFont font, FontKey id, int rank, int firstCodePoint)
            {
                Font = font;
                Id = id;
                Rank = rank;
                FirstCodePoint = firstCodePoint;
            }

            internal OpenTypeFont Font { get; private set; }
            internal FontKey Id { get; private set; }
            internal int Rank { get; private set; }
            internal int FirstCodePoint { get; private set; }
            internal HashSet<int> CodePoints { get; } = new HashSet<int>();

            /// <summary>The scripts of the code points routed to the font, in the order first met.</summary>
            internal List<UnicodeScript> Scripts { get; } = new List<UnicodeScript>();
        }

        private sealed class Assignment
        {
            internal List<Slot> Slots { get; } = new List<Slot>();
            internal List<FontKey> Order { get; } = new List<FontKey>();
            internal Dictionary<FontKey, OpenTypeFont> FontById { get; } = new Dictionary<FontKey, OpenTypeFont>();
            internal Dictionary<FontKey, HashSet<int>> CodePointsById { get; } = new Dictionary<FontKey, HashSet<int>>();

            /// <summary>Adds code points to a font's subset. A font without code points gets no subset.</summary>
            internal void Add(OpenTypeFont font, HashSet<int> codePoints)
            {
                if (codePoints == null || codePoints.Count == 0)
                    return;

                var id = IdentityOf(font);
                HashSet<int> set;
                if (!CodePointsById.TryGetValue(id, out set))
                {
                    set = new HashSet<int>();
                    CodePointsById.Add(id, set);
                    FontById.Add(id, font);
                    Order.Add(id);
                }
                set.UnionWith(codePoints);
            }
        }

        private sealed class Slot
        {
            internal Slot(Request request)
            {
                Request = request;
            }

            internal Request Request { get; private set; }
            internal FontKey? PrimaryFace { get; set; }
            internal string PrimaryName { get; set; }
            internal OpenTypeFont SafetyNet { get; set; }
            internal List<StackEntry> Fallbacks { get; } = new List<StackEntry>();
        }

        private sealed class StackEntry
        {
            internal StackEntry(Fallback source, FontKey? face, string name)
            {
                Source = source;
                Face = face;
                Name = name;
            }

            /// <summary>The fallback the entry is for.</summary>
            internal Fallback Source { get; private set; }

            internal FontKey? Face { get; private set; }
            internal string Name { get; private set; }
        }
    }
}