using Export.ApplicationService;
using Export.ApplicationService.Core.Interface;
using Microsoft.AspNetCore.Mvc;

namespace Export.Sample.API.Controllers;

[ApiController]
[Route("[controller]")]
public class WeatherForecastController : ControllerBase
{
    private static readonly string[] Summaries = new[]
    {
    "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
};

    private readonly ILogger<WeatherForecastController> _logger;
    private readonly IExportService _exportService;

    public WeatherForecastController(ILogger<WeatherForecastController> logger, IExportService exportService)
    {
        _logger = logger;
        _exportService = exportService;
    }

    [HttpGet(Name = "GetWeatherForecast")]
    public IEnumerable<WeatherForecast> Get()
    {
        return Enumerable.Range(1, 5).Select(index => new WeatherForecast
        {
            Date = DateTime.Now.AddDays(index),
            TemperatureC = Random.Shared.Next(-20, 55),
            Summary = Summaries[Random.Shared.Next(Summaries.Length)]
        })
        .ToArray();
    }

    [HttpPost("ef/csv")]
    public async Task<FileStreamResult> GetCsvFromEF()
    {
        List<WeatherForecast> weatherForecasts = CreateWeatherForecast();
        var filePath = await _exportService.ExportAsync(ApplicationService.Core.ExportType.Csv, () => weatherForecasts.AsAsyncQueryable());
        return ExportFileExtension.GetFileStreamResult(filePath);
    }

    [HttpPost("ef/xlsx")]
    public async Task<FileStreamResult> GetXlsxFromEF()
    {
        List<WeatherForecast> weatherForecasts = CreateWeatherForecast();
        var filePath = await _exportService.ExportAsync(ApplicationService.Core.ExportType.Xlsx, () => weatherForecasts.AsAsyncQueryable());
        return ExportFileExtension.GetFileStreamResult(filePath);
    }

    [HttpPost("ef/text")]
    public async Task<FileStreamResult> GetTextFromEF()
    {
        List<WeatherForecast> weatherForecasts = CreateWeatherForecast();
        var filePath = await _exportService.ExportAsync(ApplicationService.Core.ExportType.Text, () => weatherForecasts.AsAsyncQueryable());
        return ExportFileExtension.GetFileStreamResult(filePath);
    }

    [HttpPost("csv")]
    public async Task<FileStreamResult> GetCsv()
    {
        List<WeatherForecast> weatherForecasts = CreateWeatherForecast();
        var filePath = await _exportService.ExportAsync(ApplicationService.Core.ExportType.Csv, weatherForecasts);
        return ExportFileExtension.GetFileStreamResult(filePath);
    }

    [HttpPost("xlsx")]
    public async Task<FileStreamResult> GetXlsx()
    {
        List<WeatherForecast> weatherForecasts = CreateWeatherForecast();
        var filePath = await _exportService.ExportAsync(ApplicationService.Core.ExportType.Xlsx, weatherForecasts);
        return ExportFileExtension.GetFileStreamResult(filePath);
    }

    [HttpPost("text")]
    public async Task<FileStreamResult> GetText()
    {
        List<WeatherForecast> weatherForecasts = CreateWeatherForecast();
        var filePath = await _exportService.ExportAsync(ApplicationService.Core.ExportType.Text, weatherForecasts);
        return ExportFileExtension.GetFileStreamResult(filePath);
    }

   

    private List<WeatherForecast> CreateWeatherForecast()
    {
        List<WeatherForecast> weatherForecasts = new List<WeatherForecast>();
        for(int i = 0; i < 100_000; i++)
        {
            weatherForecasts.Add(new WeatherForecast
            {
                Date= DateTime.Now.AddHours(i),
                TemperatureC = i,
                Summary = $"this is summary with random guid {Guid.NewGuid()}"
            });
        }
        return weatherForecasts;
    }
}