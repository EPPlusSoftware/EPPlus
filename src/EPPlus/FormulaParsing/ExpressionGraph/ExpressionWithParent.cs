namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    /// <summary>
    /// An expression that has a parent expression.
    /// </summary>
    public abstract class ExpressionWithParent : Expression
    {
        internal ExpressionWithParent(ParsingContext ctx) : base(ctx)
        {

        }
        internal ExpressionWithParent(string expression, ParsingContext ctx) : base(expression, ctx)
        {

        }
        internal Expression _parent;
        internal int Index
        {
            get
            {
                return _parent.Children.IndexOf(this);
            }
        }
    }
}
