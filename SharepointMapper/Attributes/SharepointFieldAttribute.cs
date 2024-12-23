namespace SharepointMapper.Attributes;

/// <summary>
/// Maps property to Sharepoint field by internal name 
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class SharepointFieldAttribute : Attribute
{
    public string InternalName { get; set; }

    public SharepointFieldAttribute(string internalName)
    {
        InternalName = internalName;
    }

}