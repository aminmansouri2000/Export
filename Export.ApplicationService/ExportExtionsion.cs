namespace Export.ApplicationService;

public static class ExportExtionsion
{
    public static IServiceCollection UseExport(this IServiceCollection services, IConfiguration configuration)
    {
        services.Configure<ExportOption>(configuration.GetSection("ExportOption"));
        services.TryAddScoped<IExportService, ExportService>();
        return services;
    }
}