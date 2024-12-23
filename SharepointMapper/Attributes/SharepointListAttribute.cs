namespace SharepointMapper.Attributes;

/// <summary>
/// Maps class to Sharepoint list by Title, Guid
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class SharepointListAttribute(string title) : Attribute
{
    public string Title { get; private set; } = title;
    public Guid Id { get; private set; }

    public SharepointListAttribute(Guid id) : this(string.Empty)
    {
        Id = id;
    }
}