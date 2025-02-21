using System.Collections.ObjectModel;
using System.Globalization;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace SharepointMapper.Extensions;

public static class ListItemExtensions
{
    private const string ChildItems = "_Child_Items_";
    private const string ObjectType = "_ObjectType_";

    public static TaxonomyFieldValue? GetTaxonomyFieldValue(this ListItem item,
        string internalFieldName)
    {
        ArgumentNullException.ThrowIfNull(item);
        ArgumentNullException.ThrowIfNull(internalFieldName);

        if (!item.FieldValues.TryGetValue(internalFieldName, out var fieldValue))
            throw new ArgumentException($"The field '{internalFieldName}' does not exist.",
                nameof(internalFieldName));

        return fieldValue switch
        {
            null => null,
            TaxonomyFieldValue taxonomyFieldValue => taxonomyFieldValue,
            Dictionary<string, object> dictionary => ConvertDictionaryToTaxonomyFieldValue(dictionary),
            _ => throw new InvalidOperationException(
                $"Could not convert value of field '{internalFieldName}' to a taxonomy field value. Value is neither a TaxonomyFieldValue nor a Dictionary")
        };
    }

    public static ReadOnlyCollection<TaxonomyFieldValue>? GetTaxonomyFieldValueCollection(
        this ListItem item, string internalFieldName)
    {
        ArgumentNullException.ThrowIfNull(item);
        ArgumentNullException.ThrowIfNull(internalFieldName);

        if (!item.FieldValues.TryGetValue(internalFieldName, out var fieldValue))
            throw new ArgumentException($"The field '{internalFieldName}' does not exist.",
                nameof(internalFieldName));

        return fieldValue switch
        {
            null => null,
            TaxonomyFieldValueCollection taxonomyFieldValueCollection => new ReadOnlyCollection<TaxonomyFieldValue>(
                taxonomyFieldValueCollection.ToList()),
            Dictionary<string, object> dictionary => new ReadOnlyCollection<TaxonomyFieldValue>(
                ConvertDictionaryToTaxonomyFieldValueCollection(dictionary)),
            _ => throw new InvalidOperationException(
                $"Could not convert value of field '{internalFieldName}' to a taxonomy field value. Value is neither a TaxonomyFieldValue nor a Dictionary")
        };
    }

    private static TaxonomyFieldValue ConvertDictionaryToTaxonomyFieldValue(Dictionary<string, object> dictionary)
    {
        if (!dictionary.ContainsKey(ObjectType) || !dictionary[ObjectType].Equals("SP.Taxonomy.TaxonomyFieldValue"))
            throw new InvalidOperationException("Dictionary value represents no TaxonomyFieldValue.");

        return new TaxonomyFieldValue
        {
            Label = dictionary["Label"].ToString(),
            TermGuid = dictionary["TermGuid"].ToString(),
            WssId = int.Parse(dictionary["WssId"].ToString() ?? string.Empty, CultureInfo.InvariantCulture)
        };
    }

    private static List<TaxonomyFieldValue> ConvertDictionaryToTaxonomyFieldValueCollection(
        Dictionary<string, object> dictionary)
    {
        if (!dictionary.ContainsKey(ObjectType) ||
            !dictionary[ObjectType].Equals("SP.Taxonomy.TaxonomyFieldValueCollection"))
            throw new InvalidOperationException("Dictionary value represents no TaxonomyFieldValueCollection.");

        if (!dictionary.ContainsKey(ChildItems))
            throw new InvalidOperationException(
                $"Missing '{ChildItems}' key in TaxonomyFieldValueCollection field.");

        var list = new List<TaxonomyFieldValue>();
        foreach (var value in (object[])dictionary[ChildItems])
        {
            var childDictionary = (Dictionary<string, object>)value;
            list.Add(ConvertDictionaryToTaxonomyFieldValue(childDictionary));
        }

        return list;
    }
}