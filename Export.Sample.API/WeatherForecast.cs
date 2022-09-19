using Export.ApplicationService.Core;
using Export.ApplicationService.Core.Interface;

namespace Export.Sample.API;

public class WeatherForecast : IExportableResponse
{
    public DateTime Date { get; set; }

    public int TemperatureC { get; set; }

    public int TemperatureF => 32 + (int)(TemperatureC / 0.5556);

    public string? Summary { get; set; }

    public string NumberStartZero => "00123456";

    public List<IExportCellValue> GetCellValues(int rowNumber)
    {
        return new List<IExportCellValue>
       {
           new ExportCellValueLong(rowNumber),
           new ExportCellValueDateTime(Date),
           new ExportCellValueLong(TemperatureC),
           new ExportCellValueLong(TemperatureF),
           new ExportCellValueString(Summary),
           new ExportCellValueNumberString(NumberStartZero),
           new ExportCustomCellValue(Date)
       };
    }

    public List<string> GetHeaders()
    {
        return new List<string>
        {
            "rowNumber",
            "Date",
            "TemperatureC",
            "TemperatureF",
            "Summary",
            "NumberStartZero",
            "CustomCellValue"
        };
    }
}

public class ExportCustomCellValue : ExportCellValue<DateTime>
{
    public ExportCustomCellValue(DateTime value) : base(value)
    {
    }

    public override string GetValue()
    {
        return $"Day of year is {Value.DayOfYear}";  
    }
}