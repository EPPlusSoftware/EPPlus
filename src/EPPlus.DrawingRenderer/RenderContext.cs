using EPPlus.Fonts.OpenType;
using OfficeOpenXml.Interfaces.Fonts;

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
        private readonly object _lock;
        private readonly Func<OpenTypeFontEngine> _engineFactory;
        private readonly RenderContext _root;
        private volatile OpenTypeFontEngine _fontEngine;

        public RenderContext(Func<OpenTypeFontEngine> engineFactory)
        {
            if (engineFactory == null)
                throw new ArgumentNullException("engineFactory");
            _engineFactory = engineFactory;
            _lock = new object();
            Target = FontRenderTarget.Document;
        }

        /// <summary>
        /// Creates a view of the root context for another render target. Shares the root's font engine.
        /// </summary>
        private RenderContext(RenderContext root, FontRenderTarget target)
        {
            _root = root;
            Target = target;
        }

        /// <summary>
        /// The output kind text is laid out for. Decides whether web font substitution applies.
        /// </summary>
        public FontRenderTarget Target { get; private set; }

        /// <summary>
        /// Returns a context for the given target that shares this context's font engine and cache.
        /// </summary>
        public RenderContext ForTarget(FontRenderTarget target)
        {
            if (target == Target)
                return this;

            var root = _root ?? this;
            return target == root.Target ? root : new RenderContext(root, target);
        }

        public OpenTypeFontEngine FontEngine
        {
            get
            {
                if (_root != null)
                    return _root.FontEngine;

                if (_fontEngine == null)
                {
                    lock (_lock)
                    {
                        if (_fontEngine == null)
                            _fontEngine = _engineFactory();
                    }
                }
                return _fontEngine;
            }
        }

        /// <summary>
        /// Returns the font family to lay out and output text with for this context's target.
        /// </summary>
        public string GetFamilyForTarget(string fontName)
        {
            return FontEngine.GetFamilyForTarget(fontName, Target);
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