using System.Linq.Expressions;

namespace SharepointMapper;

//TODO:
// Converts lambda expression to camlex filter: Func<T, bool> => Func<ListItem, bool> (where T : ISharepointItem)
public class ExpressionConverter : ExpressionVisitor
{
    
}