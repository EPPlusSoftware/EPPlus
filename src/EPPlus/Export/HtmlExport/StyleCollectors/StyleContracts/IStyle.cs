

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts
{
    public interface IStyle
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
