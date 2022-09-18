namespace Export.ApplicationService;

public static class ExportExtionsion
{
    public static IServiceCollection UseExport(this IServiceCollection services, IConfiguration configuration)
    {
        services.Configure<ExportOption>(configuration.GetSection("ExportOption"));
        services.TryAddScoped<IExportService, ExportService>();
        services.TryAddScoped<IExportToFileServiceFactory, ExportToFileServiceFactory>();
        return services;
    }
}