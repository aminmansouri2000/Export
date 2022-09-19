namespace Export.ApplicationService.Infra;

internal class ExportToCsvService : IExportToFileService
{
    private FileStream _fileStream;
    private ZipArchive _archive;
    private Stream _entryStream;
    private StreamWriter _streamWriter;

    public async Task WriteCellValues<TSearchResponse>(int startRowIndex,
        List<TSearchResponse> results)
        where TSearchResponse : IExportableResponse
    {
        StringBuilder builder = new StringBuilder();
        foreach (TSearchResponse exportableResponse in results)
        {
            List<IExportCellValue> cellValues = exportableResponse.GetCellValues(startRowIndex++);
            foreach (IExportCellValue cellValue in cellValues)
            {
                AppendBuilder(builder, (dynamic)cellValue);
            }
            builder.AppendLine();
        }

        await _streamWriter.WriteAsync(builder.ToString());
        await _streamWriter.FlushAsync();
        builder.Clear();
    }

    public async Task WriteHeaders(List<string> headers)
    {
        await _streamWriter.WriteLineAsync(string.Join(",", headers));
        await _streamWriter.FlushAsync();
    }

    public Task<string> OpenFile(string folder, string fileName)
    {
        string fullFilePath = Path.Combine(folder, $"{fileName}.zip");
        _fileStream = new FileStream(fullFilePath, FileMode.CreateNew);
        _archive = new ZipArchive(_fileStream, ZipArchiveMode.Create, true);
        ZipArchiveEntry demoFile = _archive.CreateEntry($"{fileName}.csv", CompressionLevel.Optimal);
        _entryStream = demoFile.Open();
        _streamWriter = new StreamWriter(_entryStream, Encoding.UTF8);
        return Task.FromResult(fullFilePath);
    }

    public async Task CloseFile()
    {
        await _streamWriter.DisposeAsync();
        await _entryStream.DisposeAsync();
        _archive.Dispose();
        await _fileStream.DisposeAsync();
    }

    private void AppendBuilder(StringBuilder builder, ExportCellValueString cellValue)
    {
        var addStartEndQut = false;
        string value = cellValue?.Value ?? string.Empty;
        value = value.Replace("\t", " ")
            .Replace("\r", " ")
            .Replace("\n", " ");
        if (value.Contains("\""))
        {
            value = value.Replace("\"", "\"\"");
            addStartEndQut = true;
        }

        if (value.Contains(","))
        {
            addStartEndQut = true;
        }

        if (addStartEndQut)
        {
            value = $"\"{value}\"";
        }

        builder.Append($"{value.Trim()},");
    }

    private void AppendBuilder(StringBuilder builder, ExportCellValueNumberString cellValue)
    {
        builder.Append($"=\"{cellValue?.Value}\",");
    }

    private void AppendBuilder(StringBuilder builder, ExportCellValueLong cellValue)
    {
        builder.Append($"=\"{cellValue?.Value}\",");
    }

    private void AppendBuilder(StringBuilder builder, ExportCellValueDecimal cellValue)
    {
        builder.Append($"=\"{cellValue?.Value}\",");
    }

    private void AppendBuilder(StringBuilder builder, ExportCellValueDateTime cellValue)
    {
        builder.Append($"{cellValue?.GetValue()},");
    }

    private void AppendBuilder(StringBuilder builder, IExportCellValue cellValue)
    {
        builder.Append($"{cellValue.GetValue()},");
    }
}