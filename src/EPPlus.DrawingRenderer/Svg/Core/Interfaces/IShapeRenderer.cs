using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.DrawingRenderer
{
    public interface IShapeRenderer<T>
    {
        IBasicIShapesRenderer<T> BasicShapesRenderer { get; }
        T OutputStream { get; }
        bool PreRender(List<RenderItem> items);
        BoundingBox Bounds { get; }
        string ViewBox { get; set; }
        bool Render(List<RenderItem> items);
    }
}
