using EPPlus.Fonts.OpenType;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText;

namespace EPPlus.DrawingRenderer
{
    /// <summary>
    /// Carries rendering-wide resources down the drawing render stack (DrawingRenderer and
    /// below), independent of output format (SVG, PDF). Owned by the workbook, created once
    /// per workbook. The font engine is lazy-loaded on first use so constructing the context
    /// is cheap; the expensive engine (and its font cache) is only built when something is
    /// actually rendered.
    /// </summary>
    public class RenderContext : IDisposable
    {
        // A context is either a root, which owns the engine factory and creates the engine, or a
        // view, which has a root and delegates every engine access to it. Exactly one of
        // _engineFactory and _root is set.
        private readonly object _lock = new object();
        private readonly Func<OpenTypeFontEngine>? _engineFactory;
        private readonly RenderContext? _root;
        private volatile OpenTypeFontEngine? _fontEngine;

        public RenderContext(Func<OpenTypeFontEngine> engineFactory)
        {
            if (engineFactory == null)
                throw new ArgumentNullException("engineFactory");
            _engineFactory = engineFactory;
            Target = FontRenderTarget.Document;
        }

        /// <summary>
        /// Creates a view of the root context for another render target. Shares the root's font engine.
        /// </summary>
        private RenderContext(RenderContext root, FontRenderTarget target, bool substituteWebFonts)
        {
            if (root == null)
                throw new ArgumentNullException("root");
            _root = root;
            Target = target;
            SubstituteWebFonts = substituteWebFonts;
        }

        /// <summary>
        /// The output kind text is laid out for.
        /// </summary>
        public FontRenderTarget Target { get; private set; }

        /// <summary>
        /// True when fonts are substituted for the web before text is measured, so text is laid out
        /// with a font a browser is likely to have (e.g. Aptos Narrow is measured as Calibri).
        /// False when the requested font is measured as is, for instance because the fonts are
        /// embedded in the output. Substitution only ever applies to <see cref="FontRenderTarget.Web"/>.
        /// </summary>
        public bool SubstituteWebFonts { get; private set; }

        /// <summary>
        /// Returns a context for the given target that shares this context's font engine and cache.
        /// Web font substitution applies when the target is <see cref="FontRenderTarget.Web"/>.
        /// </summary>
        public RenderContext ForTarget(FontRenderTarget target)
        {
            return ForTarget(target, target == FontRenderTarget.Web);
        }

        /// <summary>
        /// Returns a context for the given target that shares this context's font engine and cache.
        /// </summary>
        /// <param name="target">The output kind text is laid out for.</param>
        /// <param name="substituteWebFonts">
        /// Whether fonts are substituted for the web before measuring. Has no effect unless
        /// <paramref name="target"/> is <see cref="FontRenderTarget.Web"/>.
        /// </param>
        public RenderContext ForTarget(FontRenderTarget target, bool substituteWebFonts)
        {
            if (target == Target && substituteWebFonts == SubstituteWebFonts)
                return this;

            var root = _root ?? this;
            if (target == root.Target && substituteWebFonts == root.SubstituteWebFonts)
                return root;

            return new RenderContext(root, target, substituteWebFonts);
        }

        /// <summary>
        /// The font engine, created on first use by the root context. A view returns its root's engine.
        /// </summary>
        public OpenTypeFontEngine FontEngine
        {
            get
            {
                if (_root != null)
                    return _root.FontEngine;

                var engine = _fontEngine;
                if (engine != null)
                    return engine;

                lock (_lock)
                {
                    if (_fontEngine == null)
                    {
                        if (_engineFactory == null)
                            throw new InvalidOperationException("The render context has neither a root context nor an engine factory.");

                        var created = _engineFactory();
                        if (created == null)
                            throw new InvalidOperationException("The engine factory returned null.");
                        _fontEngine = created;
                    }
                    return _fontEngine;
                }
            }
        }

        /// <summary>
        /// Returns the font family to lay out and output text with for this context. The family is
        /// substituted only when <see cref="SubstituteWebFonts"/> is true and the target is
        /// <see cref="FontRenderTarget.Web"/>; otherwise it is returned unchanged.
        /// </summary>
        public string GetFamilyForTarget(string fontName)
        {
            if (!SubstituteWebFonts)
                return fontName;
            return FontEngine.GetFamilyForTarget(fontName, Target);
        }

        /// <summary>
        /// Gets a measurement shaper for the font, with the family substituted for this context's target.
        /// All text measurement in drawing rendering should go through here.
        /// </summary>
        public ITextShaper GetMeasurementShaper(IFontFormatBase font)
        {
            if (font == null)
                throw new ArgumentNullException("font");
            return FontEngine.GetMeasurementShaper(GetFamilyForTarget(font.Family), font.SubFamily);
        }

        public void Dispose()
        {
            // A view does not own the engine.
            if (_root != null)
                return;

            var engine = _fontEngine;
            if (engine != null)
            {
                try { engine.Dispose(); } catch { /* best effort */ }
                _fontEngine = null;
            }
        }
    }
}