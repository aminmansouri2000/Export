namespace Export.ApplicationService.Core.Interface;

public interface IExportToFileService
{
    Task WriteCellValues<TSearchResponse>(int startRowIndex,
        List<TSearchResponse> results)
        where TSearchResponse : IExportableResponse;

    Task WriteHeaders(List<string> headers);

    Task<string> OpenFile(string folder, string fileName);

    Task CloseFile();
}