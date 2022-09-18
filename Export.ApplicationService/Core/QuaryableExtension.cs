namespace Export.ApplicationService.Core;

internal static class QuaryableExtension
{
    internal static IQueryable<T> ApplyPagingFilter<T>(this IQueryable<T> query, int pageNumber, int pageSize)
    {
        int skipSize = pageSize * (pageNumber - 1);
        if(skipSize > 0)
            query = query.Skip(skipSize);
        query = query.Take(pageSize);   
        return query;
    }
}
