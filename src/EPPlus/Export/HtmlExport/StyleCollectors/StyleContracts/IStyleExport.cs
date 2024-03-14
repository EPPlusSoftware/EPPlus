

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts
{
    public interface IStyleExport
    {
        internal string StyleKey { get; }

        internal bool HasStyle { get; }

        /// <summary>
        /// Fill
        /// </summary>
        internal IFill Fill { get; }

        /// <summary>
        /// Font
        /// </summary>
        internal IFont Font { get; }

        /// <summary>
        /// Border
        /// </summary>
        internal IBorder Border { get; }
    }
}
