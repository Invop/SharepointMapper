Is a lightweight mapper for Sharepoint online lists (CSOM).

The basic idea belongs to :"https://github.com/fixer-m/SharepointMapper"

## Mapping Guide
1. Your entity class should implement **ISharepointItem**. 
2. Use attribute **[SharepointList]** to map your entity class to Sharepoint list. Set list **name or guid**. 
3. Use attribute **[SharepointField]** to map your property to Sharepoint item field. Set field **Internal Name**. 

```csharp
[SharepointList("Customers")]
public class SpProduct : ISharepointItem
{
    [SharepointField("ID")]
    public int Id { get; set; }

    [SharepointField("CustomerTitleEn")]
    public string Title { get; set; }

    [SharepointField("RankStatus")]
    public int? Rank { get; set; }

    [SharepointField("CustomerCreated")]
    public DateTime? Created { get; set; }

    [SharepointField("NumericInfo")]
    public double? SomeNumeric { get; set; }

    [SharepointField("YesNoColumn")]
    public bool? YesNoInfo { get; set; }

    [SharepointField("TaxonomyField")]
    public TaxonomyFieldValue Field1 { get; set; }

    [SharepointField("TaxonomyFieldCollection")]
    public ReadOnlyCollection<TaxonomyFieldValue> Field2 { get; set; }

    [SharepointField("LookupField")]
    public FieldLookupValue LookupField { get; set; }
}
```
Log:
04.02.2025 :
1) Modified SharePointListAttribute to parse input as either a Title or a GUID