namespace Export.ApplicationService.Core.Interface;

public interface IExportService
{
    Task<string> ExportAsync<TSearchResponse>(ExportType exportType, Func<IQueryable<TSearchResponse>> queryFunc)
        where TSearchResponse : IExportableResponse;

    Task<string> ExportAsync<TSearchResponse>(ExportType exportType, IEnumerable<TSearchResponse> responses)
        where TSearchResponse : IExportableResponse;

}