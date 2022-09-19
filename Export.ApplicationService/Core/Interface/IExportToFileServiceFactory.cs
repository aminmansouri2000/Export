namespace Export.ApplicationService.Core.Interface;

internal interface IExportToFileServiceFactory
{
    IExportToFileService GetExportService(ExportType exportType);
}