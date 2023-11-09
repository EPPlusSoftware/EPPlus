
namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts
{
    public interface IBorder
    {
        bool HasValue { get; }

        IBorderItem Top { get; }
        IBorderItem Bottom { get; }
        IBorderItem Left { get; }
        IBorderItem Right { get; }
    }
}
