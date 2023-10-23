

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts
{
    public interface IStyleExport
    {
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
