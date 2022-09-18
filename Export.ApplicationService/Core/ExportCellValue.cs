namespace Export.ApplicationService.Core;

public abstract class ExportCellValue<T> : IExportCellValue
{
    public T Value { get; }

    protected ExportCellValue(T value)
    {
        Value = value;
    }

    public virtual string GetValue()
    {
        return Value?.ToString();
    }
}

public class ExportCellValueString : ExportCellValue<string>
{
    public ExportCellValueString(string value) : base(value)
    {
    }
}

public class ExportCellValueNumberString : ExportCellValue<string>
{
    public ExportCellValueNumberString(string value) : base(value)
    {
    }
}

public class ExportCellValueLong : ExportCellValue<long>
{
    public ExportCellValueLong(long value) : base(value)
    {
    }
}

public class ExportCellValueDecimal : ExportCellValue<decimal>
{
    public ExportCellValueDecimal(decimal value) : base(value)
    {
    }
}

public class ExportCellValueDateTime : ExportCellValue<DateTime?>
{
    public ExportCellValueDateTime(DateTime? value) : base(value)
    {
    }

    public override string GetValue()
    {
        return Value?.ToString("yyyy-MM-dd HH:mm-ss");
    }
}

