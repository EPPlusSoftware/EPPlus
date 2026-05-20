namespace EPPlus.DrawingRenderer.RenderItems
{
    public class TitleRenderItem : RenderItem
    {
        public string Title { get; private set; }

        public TitleRenderItem(string titleName) : base()
        {
            Title = titleName;
        }

        public override RenderItemType Type => RenderItemType.CommentTitle;

    }
}