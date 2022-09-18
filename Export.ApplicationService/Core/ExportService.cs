using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;

namespace Export.ApplicationService.Core;

public class ExportService : IExportService
{
    private readonly ILogger<ExportService> _logger;
    private readonly IOptionsMonitor<ExportOption> _options;
    private readonly IExportToFileServiceFactory _exportServiceFactory;

    public ExportService(ILogger<ExportService> logger,
        IOptionsMonitor<ExportOption> options,
        IExportToFileServiceFactory exportServiceFactory)
    {
        _logger = logger;
        _options = options;
        _exportServiceFactory = exportServiceFactory;
    }

    public async Task<string> ExportAsync<TSearchResponse>(ExportType exportType,
        Func<IQueryable<TSearchResponse>> queryFunc)
        where TSearchResponse : IExportableResponse
    {
        string fileName = $"{typeof(TSearchResponse).Name}-{DateTime.Now.ToFileTime()}";
        int pageSize = _options.CurrentValue.MaxQueryCount;
        int pageNumber = 1;
        int rowNumber = 1;
        bool isHeaderAdded = false;

        IExportToFileService exportService = _exportServiceFactory.GetExportService(exportType);
        string fullFilePath = await exportService.OpenFile(_options.CurrentValue.OutputFolder, fileName);
        List<TSearchResponse> results;
        do
        {
            _logger.LogInformation("start export page {@pageNumber} to file.", pageNumber);

            results = await queryFunc()
                .ApplyPagingFilter(pageNumber, pageSize)
                .ToListAsync();

            if (!isHeaderAdded)
            {
                List<string> headers = results.FirstOrDefault()?.GetHeaders() ?? new List<string>();
                await exportService.WriteHeaders(headers);
                isHeaderAdded = true;
            }

            Stopwatch stopwatch = Stopwatch.StartNew();
            await exportService.WriteCellValues(rowNumber, results);
            stopwatch.Stop();
            _logger.LogInformation("WriteCellValues for Page {@pageNumber} done in {@time} ms.", pageNumber, stopwatch.ElapsedMilliseconds);

            _logger.LogInformation("end export page {@pageNumber} to file with {@count} records.", pageNumber, results.Count);
            rowNumber += results.Count;
            pageNumber++;
        } while (results.Count == pageSize);

        await exportService.CloseFile();

        return fullFilePath;
    }

    public async Task<string> ExportAsync<TSearchResponse>(ExportType exportType,
        IEnumerable<TSearchResponse> responses)
        where TSearchResponse : IExportableResponse
    {
        string fileName = $"{typeof(TSearchResponse).Name}-{DateTime.Now.ToFileTime()}";

        IExportToFileService exportService = _exportServiceFactory.GetExportService(exportType);
        string fullFilePath = await exportService.OpenFile(_options.CurrentValue.OutputFolder, fileName);
        _logger.LogInformation("start export to file.");

        List<string> headers = responses.FirstOrDefault()?.GetHeaders() ?? new List<string>();
        await exportService.WriteHeaders(headers);

        Stopwatch stopwatch = Stopwatch.StartNew();
        await exportService.WriteCellValues(1, responses.ToList());
        stopwatch.Stop();
        _logger.LogInformation("export to file finished in {@time} ms.", stopwatch.ElapsedMilliseconds);

        await exportService.CloseFile();
        return fullFilePath;
    }
}