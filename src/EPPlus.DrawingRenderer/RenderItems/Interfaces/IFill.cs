using EPPlus.DrawingRenderer;
namespace EPPlus.Export.ImageRenderer.RenderItems.Interfaces
{
    public interface IFill
    {
        string Color { get; set; }
        double? Opacity { get; set; }
        PathFillMode FillMode { get; set; }
    }
}
