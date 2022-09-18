namespace Export.ApplicationService.Core.Interface;

public interface IExportableResponse
{
    public List<IExportCellValue> GetCellValues(int rowNumber);

    public List<string> GetHeaders();
}
