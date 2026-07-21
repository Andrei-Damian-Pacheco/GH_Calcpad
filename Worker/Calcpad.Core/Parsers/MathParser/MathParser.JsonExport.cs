namespace Calcpad.Core
{
    public partial class MathParser
    {
        // Added for GH_Calcpad's JSON worker: lets ExpressionParser capture the
        // {name, value, unit} of the line just computed by Calculate(), without
        // scraping the rendered HTML. _rpn/_result are private to this class, so
        // this has to live here rather than in ExpressionParser or an external project.
        internal bool TryGetLastAssignment(out string name, out double value, out string unit)
        {
            if (_rpn is { Length: > 0 } &&
                _rpn[0].Type == TokenTypes.Variable &&
                IsAssignment(_rpn[^1].Content))
            {
                name = _rpn[0].Content;
                value = Real;
                unit = Units?.Text ?? string.Empty;
                return true;
            }
            name = null;
            value = double.NaN;
            unit = string.Empty;
            return false;
        }
    }
}
