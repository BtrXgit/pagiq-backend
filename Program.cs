using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using System.IO;
using System;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

string syncfusionKey = Environment.GetEnvironmentVariable("SYNCFUSION_LICENSE_KEY") ?? "";
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(syncfusionKey);

app.Use(async (context, next) =>
{
    var expectedToken = Environment.GetEnvironmentVariable("API_SECRET_TOKEN");
    if (!context.Request.Headers.TryGetValue("Authorization", out var extractedToken))
    {
        context.Response.StatusCode = 401;
        await context.Response.WriteAsync("Unauthorized: No Token Provided");
        return;
    }
    var token = extractedToken.ToString().Replace("Bearer ", "");
    if (token != expectedToken)
    {
        context.Response.StatusCode = 403;
        await context.Response.WriteAsync("Forbidden: Invalid Token");
        return;
    }
    await next(context);
});

// Helper for Auto-Delete Logic
async Task<IResult> ConvertFileAsync(HttpRequest request, Func<Stream, Stream, Task> convertAction)
{
    if (!request.HasFormContentType || request.Form.Files.Count == 0)
        return Results.BadRequest("No file uploaded.");

    var file = request.Form.Files[0];

    try
    {
        // Use a MemoryStream to ensure it is fully buffered and seekable
        using var inputStream = new MemoryStream();
        await file.CopyToAsync(inputStream);
        inputStream.Position = 0;

        using var outputStream = new MemoryStream();
        await convertAction(inputStream, outputStream);
        outputStream.Position = 0;

        return Results.File(outputStream.ToArray(), "application/pdf", "converted.pdf");
    }
    catch (Exception ex)
    {
        return Results.Problem(ex.Message);
    }
}

// ENDPOINTS
app.MapPost("/convert/word-to-pdf", (HttpRequest request) => ConvertFileAsync(request, async (input, output) =>
{
    using (WordDocument wordDocument = new WordDocument(input, Syncfusion.DocIO.FormatType.Automatic))
    using (DocIORenderer render = new DocIORenderer())
    using (PdfDocument pdfDocument = render.ConvertToPDF(wordDocument))
    {
        pdfDocument.Save(output);
    }
    await Task.CompletedTask;
}));

app.MapPost("/convert/excel-to-pdf", (HttpRequest request) => ConvertFileAsync(request, async (input, output) =>
{
    using (ExcelEngine excelEngine = new ExcelEngine())
    {
        IApplication application = excelEngine.Excel;
        application.DefaultVersion = ExcelVersion.Excel2016;
        IWorkbook workbook = application.Workbooks.Open(input);
        XlsIORenderer renderer = new XlsIORenderer();
        using (PdfDocument pdfDocument = renderer.ConvertToPDF(workbook))
        {
            pdfDocument.Save(output);
        }
        workbook.Close();
    }
    await Task.CompletedTask;
}));

app.MapPost("/convert/ppt-to-pdf", (HttpRequest request) => ConvertFileAsync(request, async (input, output) =>
{
    using (IPresentation presentation = Presentation.Open(input))
    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(presentation))
    {
        pdfDocument.Save(output);
    }
    await Task.CompletedTask;
}));

app.MapPost("/convert/rtf-to-pdf", (HttpRequest request) => ConvertFileAsync(request, async (input, output) =>
{
    using (WordDocument wordDocument = new WordDocument(input, Syncfusion.DocIO.FormatType.Rtf))
    using (DocIORenderer render = new DocIORenderer())
    using (PdfDocument pdfDocument = render.ConvertToPDF(wordDocument))
    {
        pdfDocument.Save(output);
    }
    await Task.CompletedTask;
}));

app.MapPost("/convert/odt-to-pdf", (HttpRequest request) => ConvertFileAsync(request, async (input, output) =>
{
    using (WordDocument wordDocument = new WordDocument(input, Syncfusion.DocIO.FormatType.Odt))
    using (DocIORenderer render = new DocIORenderer())
    using (PdfDocument pdfDocument = render.ConvertToPDF(wordDocument))
    {
        pdfDocument.Save(output);
    }
    await Task.CompletedTask;
}));

// --- PYTHON POWERED ENDPOINTS ---

app.MapPost("/convert/pdf-to-word", async (HttpRequest request) =>
{
    if (!request.HasFormContentType || request.Form.Files.Count == 0) return Results.BadRequest("No file uploaded.");
    var file = request.Form.Files[0];
    var tempInput = Path.GetTempFileName() + ".pdf";
    var tempOutput = Path.GetTempFileName() + ".docx";
    var scriptPath = Path.GetTempFileName() + ".py";

    try
    {
        using (var stream = new FileStream(tempInput, FileMode.Create)) await file.CopyToAsync(stream);
        
        string pyScript = $@"
from pdf2docx import Converter
cv = Converter('{tempInput}')
cv.convert('{tempOutput}')
cv.close()
";
        await File.WriteAllTextAsync(scriptPath, pyScript);

        var process = new System.Diagnostics.Process { StartInfo = new System.Diagnostics.ProcessStartInfo { FileName = "python3", Arguments = scriptPath, RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, CreateNoWindow = true } };
        process.Start(); await process.WaitForExitAsync();

        if (!File.Exists(tempOutput)) return Results.Problem("Python failed: " + await process.StandardError.ReadToEndAsync());
        return Results.File(await File.ReadAllBytesAsync(tempOutput), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "converted.docx");
    }
    finally { File.Delete(tempInput); File.Delete(tempOutput); File.Delete(scriptPath); }
});

app.MapPost("/convert/pdf-to-excel", async (HttpRequest request) =>
{
    if (!request.HasFormContentType || request.Form.Files.Count == 0) return Results.BadRequest("No file uploaded.");
    var file = request.Form.Files[0];
    var tempInput = Path.GetTempFileName() + ".pdf";
    var tempOutput = Path.GetTempFileName() + ".xlsx";
    var scriptPath = Path.GetTempFileName() + ".py";

    try
    {
        using (var stream = new FileStream(tempInput, FileMode.Create)) await file.CopyToAsync(stream);
        
        string pyScript = $@"
import pdfplumber
import pandas as pd
with pdfplumber.open('{tempInput}') as pdf:
    all_dfs = []
    for page in pdf.pages:
        tables = page.extract_tables()
        for table in tables:
            if table and len(table) > 1: # Basic check for valid table
                df = pd.DataFrame(table[1:], columns=table[0])
                all_dfs.append(df)
    if all_dfs:
        with pd.ExcelWriter('{tempOutput}') as writer:
            for i, df in enumerate(all_dfs):
                df.to_excel(writer, sheet_name=f'Table_{{i+1}}', index=False)
    else:
        df = pd.DataFrame(['No tables found in this PDF'])
        df.to_excel('{tempOutput}', index=False)
";
        await File.WriteAllTextAsync(scriptPath, pyScript);

        var process = new System.Diagnostics.Process { StartInfo = new System.Diagnostics.ProcessStartInfo { FileName = "python3", Arguments = scriptPath, RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, CreateNoWindow = true } };
        process.Start(); await process.WaitForExitAsync();

        if (!File.Exists(tempOutput)) return Results.Problem("Python failed: " + await process.StandardError.ReadToEndAsync());
        return Results.File(await File.ReadAllBytesAsync(tempOutput), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "converted.xlsx");
    }
    finally { File.Delete(tempInput); File.Delete(tempOutput); File.Delete(scriptPath); }
});

app.MapPost("/convert/pdf-to-ppt", async (HttpRequest request) =>
{
    if (!request.HasFormContentType || request.Form.Files.Count == 0) return Results.BadRequest("No file uploaded.");
    var file = request.Form.Files[0];
    var tempInput = Path.GetTempFileName() + ".pdf";
    var tempOutput = Path.GetTempFileName() + ".pptx";
    var scriptPath = Path.GetTempFileName() + ".py";

    try
    {
        using (var stream = new FileStream(tempInput, FileMode.Create)) await file.CopyToAsync(stream);
        
        string pyScript = $@"
import pdf2image
from pptx import Presentation
from pptx.util import Inches
import os

images = pdf2image.convert_from_path('{tempInput}')
prs = Presentation()
for i, img in enumerate(images):
    img_path = f'{tempInput}_{{i}}.jpg'
    img.save(img_path, 'JPEG')
    slide = prs.slides.add_slide(prs.slide_layouts[6]) # blank layout
    slide.shapes.add_picture(img_path, Inches(0), Inches(0), width=prs.slide_width)
    os.remove(img_path)
prs.save('{tempOutput}')
";
        await File.WriteAllTextAsync(scriptPath, pyScript);

        var process = new System.Diagnostics.Process { StartInfo = new System.Diagnostics.ProcessStartInfo { FileName = "python3", Arguments = scriptPath, RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, CreateNoWindow = true } };
        process.Start(); await process.WaitForExitAsync();

        if (!File.Exists(tempOutput)) return Results.Problem("Python failed: " + await process.StandardError.ReadToEndAsync());
        return Results.File(await File.ReadAllBytesAsync(tempOutput), "application/vnd.openxmlformats-officedocument.presentationml.presentation", "converted.pptx");
    }
    finally { File.Delete(tempInput); File.Delete(tempOutput); File.Delete(scriptPath); }
});

app.Run();
