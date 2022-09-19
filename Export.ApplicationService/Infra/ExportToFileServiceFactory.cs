namespace Export.ApplicationService.Infra;

internal class ExportToFileServiceFactory : IExportToFileServiceFactory
{
    public IExportToFileService GetExportService(ExportType exportType)
    {
        switch (exportType)
        {
            case ExportType.Csv:
                return new ExportToCsvService();
            case ExportType.Xlsx:
                return new ExportToXlsxService();
            case ExportType.Text:
                return new ExportToTextService();
        }
        throw new NotSupportedException($"invalid exportType: {exportType}");
    }
}