using System.Collections.ObjectModel;
using System.Reflection;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using SharepointMapper.Attributes;
using SharepointMapper.Extensions;

namespace SharepointMapper;

internal class SharepointMapper(ClientContext context)
{
    /// <summary>
    /// Returns SharePoint List for the item (via attribute SharepointList).
    /// </summary>
    /// <typeparam name="T">The type representing the SharePoint item with the SharepointListAttribute.</typeparam>
    /// <returns>The SharePoint list associated with the type T.</returns>
    /// <exception cref="ArgumentException">
    /// Thrown when type T does not have the SharepointListAttribute, or both Title and Id properties are invalid.
    /// </exception>
    public List GetListForSharepointItem<T>()
    {
        Type spEntityType = typeof(T);
        
        var spAttr = spEntityType.GetCustomAttribute<SharepointListAttribute>();

        if (spAttr == null)
        {
            throw new ArgumentException(
                $"The type '{spEntityType.FullName}' does not have the required [SharepointList] attribute. Ensure the type is decorated with this attribute."
            );
        }

        if (!string.IsNullOrWhiteSpace(spAttr.Title))
        {
            return context.Web.Lists.GetByTitle(spAttr.Title);
        }

        if (spAttr.Id != Guid.Empty)
        {
            return context.Web.Lists.GetById(spAttr.Id);
        }

        throw new InvalidOperationException(
            $"The [SharepointList] attribute on type '{spEntityType.FullName}' is invalid. Both Title and Id are either empty or not set. " +
            "Provide at least one valid property to identify the SharePoint list."
        );
    }

    /// <summary>
    /// Returns internal names of mapped fields for type T
    /// </summary>
    public List<string> GetMappedFields<T>()
    {
        List<string> mappedFields = new List<string>();
        Type spItemType = typeof(T);
        var objProperties = spItemType.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty)
            .Where(p => p.IsDefined(typeof(SharepointFieldAttribute))).ToList();

        foreach (var objProperty in objProperties)
        {
            var spAttr = objProperty.GetCustomAttribute<SharepointFieldAttribute>();
            if (spAttr != null && !mappedFields.Contains(spAttr.InternalName))
                mappedFields.Add(spAttr.InternalName);
        }

        return mappedFields;
    }
    
    private List<string> GetWritableFieldsOfList(List list)
    {
        List<string> writableFields = (list.Fields as IEnumerable<Field>).Where(f => !f.ReadOnlyField).Select(f => f.InternalName).ToList();
        writableFields.Add("_ModerationStatus");
        return writableFields;
    }

    /// <summary>
    /// Updates SharePoint list items based on the provided entity objects.
    /// </summary>
    /// <typeparam name="T">The type representing the SharePoint item, which must implement the ISharepointItem interface.</typeparam>
    /// <param name="entities">The collection of entities containing the updated data to apply to the SharePoint list items.</param>
    /// <param name="list">The SharePoint list to be updated.</param>
    /// <exception cref="ArgumentNullException">
    /// Thrown when the entities collection or the list is null.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    /// Thrown when a property of the entity does not have a valid SharePoint field mapping or when the SharePoint field is not writable.
    /// </exception>
    public void UpdateItemsFromEntities<T>(IEnumerable<T> entities, List list) where T : ISharepointItem
    {
        List<string> writableFields = GetWritableFieldsOfList(list);
        List<PropertyInfo> entityProperties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty)
            .Where(p => p.IsDefined(typeof(SharepointFieldAttribute))).ToList();

        foreach (var itemToUpdate in entities)
        {
            var item = list.GetItemById(itemToUpdate.Id);

            foreach (var objProperty in entityProperties)
            {
                var spFieldAttr = objProperty.GetCustomAttribute<SharepointFieldAttribute>();
                if (spFieldAttr != null && writableFields.Contains(spFieldAttr.InternalName))
                    item[spFieldAttr.InternalName] = objProperty.GetValue(itemToUpdate);
            }
            item.Update();
        }
    }

    /// <summary>
    /// Creates items in a SharePoint list from a collection of entity objects.
    /// </summary>
    /// <typeparam name="T">The type of entities to be created in the SharePoint list. The type must implement ISharepointItem and have properties decorated with the SharepointFieldAttribute.</typeparam>
    /// <param name="entities">The collection of entities to create as list items in the specified SharePoint list.</param>
    /// <param name="list">The SharePoint list where the new items will be added.</param>
    /// <exception cref="ArgumentNullException">Thrown when the entities collection or list is null.</exception>
    /// <exception cref="InvalidOperationException">Thrown when a property does not have the SharepointFieldAttribute or its internal name does not match a writable field in the SharePoint list.</exception>
    public void CreateItemsFromEntities<T>(IEnumerable<T> entities, List list) where T : ISharepointItem
    {
        List<string> writableFields = GetWritableFieldsOfList(list);
        List<PropertyInfo> entityProperties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty)
            .Where(p => p.IsDefined(typeof(SharepointFieldAttribute))).ToList();

        var creatItemInfo = new ListItemCreationInformation();
        foreach (var itemToInsert in entities)
        {
            var item = list.AddItem(creatItemInfo);

            foreach (var objProperty in entityProperties)
            {
                var spFieldAttr = objProperty.GetCustomAttribute<SharepointFieldAttribute>();
                if (spFieldAttr == null || !writableFields.Contains(spFieldAttr.InternalName)) continue;
                
                var newValue = objProperty.GetValue(itemToInsert);
                item[spFieldAttr.InternalName] = newValue;
            }
            item.Update();
        }
    }

    /// <summary>
    /// Builds an instance of an entity of type T from a SharePoint ListItem by mapping its fields
    /// to properties decorated with the SharepointFieldAttribute.
    /// </summary>
    /// <typeparam name="T">The type of the entity to create, which must have properties
    /// annotated with SharepointFieldAttribute.</typeparam>
    /// <param name="item">The SharePoint ListItem containing field values to be mapped to the entity.</param>
    /// <returns>An instance of type T with properties populated from the ListItem fields.</returns>
    /// <exception cref="InvalidOperationException">
    /// Thrown if a field expected by a property annotated with <see cref="SharepointFieldAttribute"/> 
    /// is missing, or if there is an issue converting a SharePoint field value to the corresponding 
    /// property type on the entity.
    /// </exception>
    /// <exception cref="ArgumentNullException">Thrown if the provided <paramref name="item"/> is null.</exception>
    /// <remarks>
    /// Ensure that the SharePoint ListItem contains all expected fields with accurate mappings 
    /// defined via the <see cref="SharepointFieldAttribute"/>. Missing or incompatible field values 
    /// will cause an exception to be thrown.
    /// </remarks>
    public T BuildEntityFromItem<T>(ListItem item) where T : new()
    {
        ArgumentNullException.ThrowIfNull(item, nameof(item));

        var entity = new T();

        const BindingFlags bindingFlags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty;
        List<PropertyInfo> entityProperties = typeof(T).GetProperties(bindingFlags)
            .Where(p => p.IsDefined(typeof(SharepointFieldAttribute))).ToList();

        var fieldsOfItem = item.FieldValues.Select(f => f.Key).ToList();

        foreach (var entityProperty in entityProperties)
        {
            var spAttr = entityProperty.GetCustomAttribute<SharepointFieldAttribute>();
            if (spAttr == null)
            {
                throw new InvalidOperationException(
                    $"Property '{entityProperty.Name}' in type '{typeof(T).FullName}' is missing a SharepointFieldAttribute.");
            }

            var spFieldName = spAttr.InternalName;

            if (!fieldsOfItem.Contains(spFieldName))
            {
                throw new InvalidOperationException(
                    $"The field '{spFieldName}' required by property '{entityProperty.Name}' was not found in the SharePoint ListItem.");
            }

            try
            {
                var fieldValue = ResolveFieldValue(entityProperty, item, spAttr);

                if (fieldValue == null)
                {
                    // Explicitly setting null for the property if fieldValue is null
                    entityProperty.SetValue(entity, null);
                }
                else
                {
                    // Convert the field value to the appropriate type and set it
                    Type typeToConvert = GetUnderlyingType(entityProperty);
                    object settablePropertyValue = Convert.ChangeType(fieldValue, typeToConvert);
                    entityProperty.SetValue(entity, settablePropertyValue);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Failed to map SharePoint field '{spFieldName}' to property '{entityProperty.Name}' in type '{typeof(T).FullName}'.",
                    ex);
            }
        }

        return entity;
    }

    /// <summary>
    /// Retrieves the underlying type of property, accounting for nullable types.
    /// </summary>
    /// <param name="objProperty">The PropertyInfo object representing the property to retrieve the underlying type for.</param>
    /// <returns>The underlying type of the property. If the property is nullable, returns the type it wraps; otherwise, returns the property's type.</returns>
    /// <exception cref="ArgumentNullException">Thrown if the objProperty parameter is null.</exception>
    private Type GetUnderlyingType(PropertyInfo objProperty)
    {
        ArgumentNullException.ThrowIfNull(objProperty);
        return Nullable.GetUnderlyingType(objProperty.PropertyType) ?? objProperty.PropertyType;
    }

    /// <summary>
    /// Retrieves the value of a SharePoint field for a property, including handling complex types such as taxonomy fields
    /// or calculated fields with errors.
    /// </summary>
    /// <param name="property">The property of the entity being populated, annotated with SharepointFieldAttribute.</param>
    /// <param name="listItem">The SharePoint ListItem containing the field values.</param>
    /// <param name="attribute">The SharepointFieldAttribute that provides the internal name of the SharePoint field.</param>
    /// <returns>
    /// The value of the SharePoint field associated with the property, converted to the appropriate type, or null if the field value is not set.
    /// </returns>
    /// <exception cref="InvalidOperationException">
    /// Thrown when there is an error resolving the field value for the specified attribute and property.
    /// </exception>
    private object? ResolveFieldValue(PropertyInfo property, ListItem listItem,
        SharepointFieldAttribute attribute)
    {
        try
        {
            if (property.PropertyType == typeof(TaxonomyFieldValue))
                return listItem.GetTaxonomyFieldValue(attribute.InternalName);

            if (property.PropertyType == typeof(ReadOnlyCollection<TaxonomyFieldValue>))
                return listItem.GetTaxonomyFieldValueCollection(attribute.InternalName);
            
            var fieldValue = listItem.FieldValues[attribute.InternalName];
            
            if (fieldValue is FieldCalculatedErrorValue)
                fieldValue = "Calculated field contains error.";
            
            return fieldValue;
        }
        catch (Exception e)
        {
            throw new InvalidOperationException(
                $"Error resolving field value for field '{attribute.InternalName}' of type '{property.PropertyType.FullName}'. \n {e.Message}",
                e);
        }
    }
}