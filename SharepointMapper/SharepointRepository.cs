using System.Linq.Expressions;
using System.Net;
using CamlexNET;
using Microsoft.SharePoint.Client;

namespace SharepointMapper;

public class SharepointRepository
{
    private readonly ClientContext _context;
    private SharepointMapper _mapper;

    public SharepointRepository(ClientContext context)
    {
        _context = context;
        _mapper = new SharepointMapper(_context);
    }
    
    //TODO: Filter query
    /// <summary>
    /// Get items of type <T> by lambda expression - filter
    /// </summary>
    public List<T> Query<T>(Expression<Func<T, bool>> filter) where T : ISharepointItem, new()
    {
        throw new NotImplementedException();
        var camlexEpressionFilter = new ExpressionConverter().Visit(filter) as Expression<Func<ListItem, bool>>;

        List<string> fieldsToLoad = _mapper.GetMappedFields<T>();
        string spQuery = Camlex.Query().Where(camlexEpressionFilter).ViewFields(fieldsToLoad).ToString(true);
        return Query<T>(spQuery);
    }

    /// <summary>
    /// Get items of type <T> by caml query string. Empty query returns all items.
    /// </summary>
    public List<T> Query<T>(string camlQueryString) where T : ISharepointItem, new()
    {
        CamlQuery query = new CamlQuery { ViewXml = camlQueryString };
        return Query<T>(query);
    }

    /// <summary>
    /// Get all items of type <T>.
    /// </summary>
    public List<T> GetAll<T>() where T : ISharepointItem, new()
    {
        CamlQuery query = new CamlQuery();
        return Query<T>(query);
    }

    private List<T> Query<T>(CamlQuery query) where T : ISharepointItem, new()
    {
        if (query.ViewXml == null)
        {
            List<string> fieldsToLoad = _mapper.GetMappedFields<T>();
            query.ViewXml = Camlex.Query().ViewFields(fieldsToLoad).ToString(true);
        }

        var list = _mapper.GetListForSharepointItem<T>();
        var items = list.GetItems(query);
        _context.Load(items);
        _context.ExecuteQuery();

        List<T> result = new List<T>();
        foreach (ListItem item in items)
            result.Add(_mapper.BuildEntityFromItem<T>(item));

        return result;
    }
    
    
    /// <summary>
    /// Get item by ID of type <T>.
    /// </summary>
    public T GetById<T>(int id) where T : ISharepointItem, new()
    {
        var list = _mapper.GetListForSharepointItem<T>();
        var item = list.GetItemById(id);
        _context.Load(item);
        _context.ExecuteQuery();

        return _mapper.BuildEntityFromItem<T>(item);
    }

    /// <summary>
    /// Update corresponding changed item in Sharepoint.
    /// </summary>
    public void Update<T>(T entity) where T : ISharepointItem
    {
        Update((IEnumerable<T>)new[] { entity });
    }
    
    /// <summary>
    /// Update corresponding items in Sharepoint.
    /// </summary>
    public void Update<T>(IEnumerable<T> entities) where T : ISharepointItem
    {
        var list = _mapper.GetListForSharepointItem<T>();

        _context.Load(list.Fields);
        _context.ExecuteQuery();

        _mapper.UpdateItemsFromEntities(entities, list);
        _context.ExecuteQuery();
    }

    /// <summary>
    /// Remove corresponding item from Sharepoint.
    /// </summary>
    public void Delete<T>(T entity) where T : ISharepointItem
    {
        Delete((IEnumerable<T>)new[] { entity });
    }
    
    /// <summary>
    /// Remove corresponding items from Sharepoint.
    /// </summary>
    public void Delete<T>(IEnumerable<T> entities) where T : ISharepointItem
    {
        var list = _mapper.GetListForSharepointItem<T>();

        foreach (var itemToDelete in entities)
        {
            var item = list.GetItemById(itemToDelete.Id);
            item.DeleteObject();
        }
        _context.ExecuteQuery();
    }
    
    
    /// <summary>
    /// Create new item in Sharepoint.
    /// </summary>
    public void Insert<T>(T entitiy) where T : ISharepointItem
    {
        Insert((IEnumerable<T>) [entitiy]);
    }

    /// <summary>
    /// Create new items in Sharepoint.
    /// </summary>
    public void Insert<T>(IEnumerable<T> entity) where T : ISharepointItem
    {
        var list = _mapper.GetListForSharepointItem<T>();
        _context.Load(list.Fields);
        _context.ExecuteQuery();

        _mapper.CreateItemsFromEntities(entity, list);
        _context.ExecuteQuery();
    }
}