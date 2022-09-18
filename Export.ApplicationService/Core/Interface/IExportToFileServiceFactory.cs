namespace Export.ApplicationService.Core.Interface;

public interface IExportToFileServiceFactory
{
    IExportToFileService GetExportService(ExportType exportType);
}