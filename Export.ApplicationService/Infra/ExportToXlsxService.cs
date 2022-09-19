using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Export.ApplicationService.Infra;

internal class ExportToXlsxService : IExportToFileService
{
    private SpreadsheetDocument _spreadsheetDocument;
    private SheetData _sheetData;

    public Task WriteCellValues<TSearchResponse>(int startRowIndex,
        List<TSearchResponse> results)
        where TSearchResponse : IExportableResponse
    {
        foreach (TSearchResponse exportableResponse in results)
        {
            List<IExportCellValue> cellValues = exportableResponse.GetCellValues(startRowIndex++);
            Row row = new Row();
            foreach (IExportCellValue cellValue in cellValues)
            {
                row.Append(GetCell((dynamic)cellValue));
            }
            _sheetData.Append(row);
        }
        return Task.CompletedTask;
    }

    public Task WriteHeaders(List<string> headers)
    {
        Row headerRow = new Row();
        foreach (string header in headers)
        {
            Cell cell = GetCell(CellValues.String, header);
            headerRow.Append(cell);
        }
        _sheetData.Append(headerRow);
        return Task.CompletedTask;
    }

    public Task<string> OpenFile(string folder, string fileName)
    {
        string fullPath = System.IO.Path.Combine(folder, $"{fileName}.xlsx");
        _spreadsheetDocument = SpreadsheetDocument.Create(fullPath, SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart1 = _spreadsheetDocument.AddWorkbookPart();
        GenerateWorkbookPart1Content(workbookPart1);

        WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
        _sheetData = GenerateWorksheetPart1Content(worksheetPart1);

        return Task.FromResult(fullPath);
    }

    public Task CloseFile()
    {
        _spreadsheetDocument.Save();
        _spreadsheetDocument.Dispose();
        return Task.CompletedTask;
    }

    //private Cell GetCell(ExportCellValueString cellValue)
    //{
    //    return GetCell(CellValues.String, cellValue.Value);
    //}

    //private Cell GetCell(ExportCellValueNumberString cellValue)
    //{
    //    return GetCell(CellValues.String, cellValue.Value);
    //}

    //private Cell GetCell(ExportCellValueDateTimeOffset cellValue)
    //{
    //    return GetCell(CellValues.String, cellValue.GetValue());
    //}

    private Cell GetCell(ExportCellValueDecimal cellValue)
    {
        return GetCell(CellValues.Number, $"{cellValue.Value}");
    }

    private Cell GetCell(ExportCellValueLong cellValue)
    {
        return GetCell(CellValues.Number, $"{cellValue.Value}");
    }

    private Cell GetCell(IExportCellValue cellValue)
    {
        return GetCell(CellValues.String, cellValue.GetValue());
    }

    private Cell GetCell(CellValues cellValues, string value)
    {
        Cell cell = new Cell
        {
            DataType = cellValues,
        };
        CellValue children = new CellValue
        {
            Text = value
        };
        cell.Append(children);
        return cell;
    }

    private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
    {
        Workbook workbook1 = new Workbook() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x15" } };
        workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        workbook1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        workbook1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
        FileVersion fileVersion1 = new FileVersion() { ApplicationName = "xl", LastEdited = "6", LowestEdited = "6", BuildVersion = "14420" };
        WorkbookProperties workbookProperties1 = new WorkbookProperties() { DefaultThemeVersion = (UInt32Value)164011U };

        AlternateContent alternateContent1 = new AlternateContent();
        alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

        AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "x15" };

        //X15ac.AbsolutePath absolutePath1 = new X15ac.AbsolutePath() { Url = "C:\\Users\\Asus\\Downloads\\" };
        //absolutePath1.AddNamespaceDeclaration("x15ac", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac");

        //alternateContentChoice1.Append(absolutePath1);

        alternateContent1.Append(alternateContentChoice1);

        BookViews bookViews1 = new BookViews();
        WorkbookView workbookView1 = new WorkbookView() { XWindow = 0, YWindow = 0, WindowWidth = (UInt32Value)23040U, WindowHeight = (UInt32Value)9192U };

        bookViews1.Append(workbookView1);

        Sheets sheets1 = new Sheets();
        Sheet sheet1 = new Sheet() { Name = "Sheet1", SheetId = (UInt32Value)1U, Id = "rId1" };

        sheets1.Append(sheet1);
        CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)162913U };

        WorkbookExtensionList workbookExtensionList1 = new WorkbookExtensionList();

        WorkbookExtension workbookExtension1 = new WorkbookExtension() { Uri = "{140A7094-0E35-4892-8432-C4D2E57EDEB5}" };
        workbookExtension1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
        //X15.WorkbookProperties workbookProperties2 = new X15.WorkbookProperties() { ChartTrackingReferenceBase = true };

        //workbookExtension1.Append(workbookProperties2);

        workbookExtensionList1.Append(workbookExtension1);

        workbook1.Append(fileVersion1);
        workbook1.Append(workbookProperties1);
        workbook1.Append(alternateContent1);
        workbook1.Append(bookViews1);
        workbook1.Append(sheets1);
        workbook1.Append(calculationProperties1);
        workbook1.Append(workbookExtensionList1);

        workbookPart1.Workbook = workbook1;
    }

    private SheetData GenerateWorksheetPart1Content(WorksheetPart worksheetPart1)
    {
        Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
        worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
        SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:C3" };

        SheetViews sheetViews1 = new SheetViews();

        SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
        Selection selection1 = new Selection() { ActiveCell = "A9", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A9" } };

        sheetView1.Append(selection1);

        sheetViews1.Append(sheetView1);
        SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 14.4D, DyDescent = 0.3D };

        SheetData sheetData1 = new SheetData();
        PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
        PageSetup pageSetup1 = new PageSetup() { Orientation = OrientationValues.Portrait, Id = "rId1" };

        worksheet1.Append(sheetDimension1);
        worksheet1.Append(sheetViews1);
        worksheet1.Append(sheetFormatProperties1);
        worksheet1.Append(sheetData1);
        worksheet1.Append(pageMargins1);
        worksheet1.Append(pageSetup1);

        worksheetPart1.Worksheet = worksheet1;

        return sheetData1;
    }
}