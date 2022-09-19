using Microsoft.AspNetCore.Mvc;

namespace Export.ApplicationService;

public static class ExportFileExtension
{
    public static FileStreamResult GetFileStreamResult(string filePath)
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

    private static string GetFileContentType(string exportType)
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
}