namespace SharepointMapper.Attributes;

/// <summary>
/// Maps class to SharePoint list by Title or Guid.
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class SharepointListAttribute : Attribute
{
    public string Title { get; private set; } = string.Empty;
    public Guid Id { get; private set; }

    public SharepointListAttribute(string titleOrId)
    {
        if (Guid.TryParse(titleOrId, out Guid parsedId))
        {
            Id = parsedId;
        }
        else
        {
            Title = titleOrId;
        }

    }
    
}