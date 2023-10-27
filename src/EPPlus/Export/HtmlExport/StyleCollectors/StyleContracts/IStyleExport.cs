

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts
{
    public interface IStyleExport
    {
        public string StyleKey { get; }

        public bool HasStyle { get; }

        /// <summary>
        /// Fill
        /// </summary>
        IFill Fill { get; }

        /// <summary>
        /// Font
        /// </summary>
        IFont Font { get; }

        /// <summary>
        /// Border
        /// </summary>
        IBorder Border { get; }
    }
}
