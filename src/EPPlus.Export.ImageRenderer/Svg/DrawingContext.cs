using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;

namespace EPPlus.Export.ImageRenderer.Svg
{
    internal class DrawingContext<T>
    {
        public DrawingContext(T topDrawingHandler,ExcelDrawing drawing)
        {
            TopDrawingHandler = topDrawingHandler;
            Drawing = drawing;
        }        
        public T TopDrawingHandler { get; set; }
        public ExcelDrawing Drawing { get; set; }
        public ExcelTheme Theme { get; set; }
    }
}
