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
        return GetFileStreamResult(filePath);
    }

    [HttpPost("ef/xlsx")]
    public async Task<FileStreamResult> GetXlsxFromEF()
    {
        List<WeatherForecast> weatherForecasts = CreateWeatherForecast();
        var filePath = await _exportService.ExportAsync(ApplicationService.Core.ExportType.Xlsx, () => weatherForecasts.AsAsyncQueryable());
        return GetFileStreamResult(filePath);
    }

    [HttpPost("ef/text")]
    public async Task<FileStreamResult> GetTextFromEF()
    {
        List<WeatherForecast> weatherForecasts = CreateWeatherForecast();
        var filePath = await _exportService.ExportAsync(ApplicationService.Core.ExportType.Text, () => weatherForecasts.AsAsyncQueryable());
        return GetFileStreamResult(filePath);
    }

    [HttpPost("csv")]
    public async Task<FileStreamResult> GetCsv()
    {
        List<WeatherForecast> weatherForecasts = CreateWeatherForecast();
        var filePath = await _exportService.ExportAsync(ApplicationService.Core.ExportType.Csv, weatherForecasts);
        return GetFileStreamResult(filePath);
    }

    [HttpPost("xlsx")]
    public async Task<FileStreamResult> GetXlsx()
    {
        List<WeatherForecast> weatherForecasts = CreateWeatherForecast();
        var filePath = await _exportService.ExportAsync(ApplicationService.Core.ExportType.Xlsx, weatherForecasts);
        return GetFileStreamResult(filePath);
    }

    [HttpPost("text")]
    public async Task<FileStreamResult> GetText()
    {
        List<WeatherForecast> weatherForecasts = CreateWeatherForecast();
        var filePath = await _exportService.ExportAsync(ApplicationService.Core.ExportType.Text, weatherForecasts);
        return GetFileStreamResult(filePath);
    }

    private FileStreamResult GetFileStreamResult(string filePath)
    {
        if (!Path.IsPathRooted(filePath))
        {
            throw new ArgumentException($"Invalid input parameter {filePath}");
        }

        string fileExtensions = Path.GetExtension(filePath).Replace(".", "");
        FileStreamResult fileStreamResult =
            new FileStreamResult(System.IO.File.OpenRead(filePath), GetFileContentType(fileExtensions))
            {
                FileDownloadName = Path.GetFileName(filePath),
            };
        return fileStreamResult;
    }

    private string GetFileContentType(string exportType)
    {
        switch (exportType.ToLower())
        {
            case "zip": return "application/zip";
            case "xlsx":
            case "excel":
                return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            case "pdf": return "application/pdf";
            case "html": return "text/HTML";
            case "csv": return "text/HTML";
            default: return "text/plain";
        }
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