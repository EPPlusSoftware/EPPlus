using EPPlus.Fonts.OpenType;

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
        private readonly object _lock = new object();
        private readonly Func<OpenTypeFontEngine> _engineFactory;
        private OpenTypeFontEngine? _fontEngine;

        public RenderContext(Func<OpenTypeFontEngine> engineFactory)
        {
            if (engineFactory == null)
                throw new ArgumentNullException("engineFactory");
            _engineFactory = engineFactory;
        }

        public OpenTypeFontEngine FontEngine
        {
            get
            {
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


        public void Dispose()
        {
            if (_fontEngine != null)
            {
                try { _fontEngine.Dispose(); } catch { /* best effort */ }
                _fontEngine = null;
            }
        }
    }
}