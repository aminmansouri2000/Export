namespace Export.ApplicationService.Infra;

public class ExportToTextService : IExportToFileService
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
        await _streamWriter.WriteLineAsync(string.Join("\t", headers));
        await _streamWriter.FlushAsync();
    }

    public Task<string> OpenFile(string folder, string fileName)
    {
        string fullFilePath = Path.Combine(folder, $"{fileName}.zip");
        _fileStream = new FileStream(fullFilePath, FileMode.CreateNew);
        _archive = new ZipArchive(_fileStream, ZipArchiveMode.Create, true);
        ZipArchiveEntry demoFile = _archive.CreateEntry($"{fileName}.txt", CompressionLevel.Optimal);
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
        string value = cellValue?.Value ?? string.Empty;
        value = value.Replace("\t", " ")
            .Replace("\r", " ")
            .Replace("\n", " ");
        builder.Append($"{value.Trim()}\t");
    }

    private void AppendBuilder(StringBuilder builder, IExportCellValue cellValue)
    {
        builder.Append($"{cellValue.GetValue()}\t");
    }
}