using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using PuppeteerSharp;
using HtmlAgilityPack;
using MailMergeSaaS.Models;
using System.Text;
using System.Text.RegularExpressions;

namespace MailMergeSaaS.Services;

public class MailMergeService
{
    private readonly ILogger<MailMergeService> _logger;
    private readonly Dictionary<string, MailMergeSession> _sessions = new();

    public MailMergeService(ILogger<MailMergeService> logger)
    {
        _logger = logger;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        // Download Chromium for PuppeteerSharp
        _ = Task.Run(async () =>
        {
            try
            {
                var fetcher = new BrowserFetcher();
                await fetcher.DownloadAsync();
                _logger.LogInformation("Chromium downloaded for PDF generation");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to download Chromium");
            }
        });
    }

    public async Task<ProcessingResult<string>> UploadTemplateAsync(string sessionId, IFormFile template)
    {
        try
        {
            var session = GetOrCreateSession(sessionId);
            
            // Clean up previous template
            if (!string.IsNullOrEmpty(session.TemplatePath) && File.Exists(session.TemplatePath))
                File.Delete(session.TemplatePath);

            var tempDir = Path.GetTempPath();
            var fileName = $"template_{sessionId}_{DateTime.Now:yyyyMMddHHmmss}_{template.FileName}";
            var filePath = Path.Combine(tempDir, fileName);

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await template.CopyToAsync(stream);
            }

            // Validate the document can be opened
            using (var doc = WordprocessingDocument.Open(filePath, false))
            {
                if (doc.MainDocumentPart?.Document == null)
                    throw new InvalidOperationException("Invalid Word document");
            }

            session.TemplatePath = filePath;
            return ProcessingResult<string>.Success(filePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error uploading template");
            return ProcessingResult<string>.Failure($"Error uploading template: {ex.Message}");
        }
    }

    public async Task<ProcessingResult<object>> UploadDataAsync(string sessionId, IFormFile dataFile)
    {
        try
        {
            var session = GetOrCreateSession(sessionId);
            
            // Clean up previous data
            if (!string.IsNullOrEmpty(session.DataPath) && File.Exists(session.DataPath))
                File.Delete(session.DataPath);

            var tempDir = Path.GetTempPath();
            var fileName = $"data_{sessionId}_{DateTime.Now:yyyyMMddHHmmss}_{dataFile.FileName}";
            var filePath = Path.Combine(tempDir, fileName);

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await dataFile.CopyToAsync(stream);
            }

            // Load and preview Excel data
            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets.FirstOrDefault();
            
            if (worksheet == null)
                throw new InvalidOperationException("No worksheet found in Excel file");

            var data = new List<Dictionary<string, string>>();
            var headers = new List<string>();

            // Get headers from first row
            var colCount = worksheet.Dimension?.Columns ?? 0;
            for (int col = 1; col <= colCount; col++)
            {
                var header = worksheet.Cells[1, col].Text;
                headers.Add(string.IsNullOrWhiteSpace(header) ? $"Column{col}" : header);
            }

            // Get data rows (limit preview to 5 rows)
            var rowCount = Math.Min(worksheet.Dimension?.Rows ?? 0, 6);
            for (int row = 2; row <= rowCount; row++)
            {
                var rowData = new Dictionary<string, string>();
                for (int col = 1; col <= colCount; col++)
                {
                    var value = worksheet.Cells[row, col].Text;
                    rowData[headers[col - 1]] = value ?? "";
                }
                data.Add(rowData);
            }

            session.DataPath = filePath;
            session.Headers = headers;

            return ProcessingResult<object>.Success(new { headers, data });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error uploading data");
            return ProcessingResult<object>.Failure($"Error uploading data: {ex.Message}");
        }
    }

    public async Task<ProcessingResult<List<string>>> ProcessMergeAsync(string sessionId, string outputType, bool multipleFiles)
    {
        try
        {
            var session = GetSession(sessionId);
            if (session == null)
                return ProcessingResult<List<string>>.Failure("Session not found");

            if (string.IsNullOrEmpty(session.TemplatePath) || !File.Exists(session.TemplatePath))
                return ProcessingResult<List<string>>.Failure("Template file not found");

            if (string.IsNullOrEmpty(session.DataPath) || !File.Exists(session.DataPath))
                return ProcessingResult<List<string>>.Failure("Data file not found");

            // Load Excel data
            var excelData = LoadExcelData(session.DataPath);
            
            // Clean up previous output
            session.OutputFiles.Clear();

            if (multipleFiles)
            {
                // Generate separate file for each record
                for (int i = 0; i < excelData.Count; i++)
                {
                    var record = excelData[i];
                    var fileName = GetSafeFileName(record, session.Headers, i);
                    
                    var outputPath = outputType.ToLower() == "pdf" 
                        ? await GenerateSinglePdfAsync(session.TemplatePath, new List<Dictionary<string, string>> { record }, fileName)
                        : await GenerateSingleWordAsync(session.TemplatePath, new List<Dictionary<string, string>> { record }, fileName);
                    
                    if (outputPath != null)
                        session.OutputFiles.Add(Path.GetFileName(outputPath));
                }
            }
            else
            {
                // Generate single file with all records
                var fileName = $"merged_output_{DateTime.Now:yyyyMMddHHmmss}";
                
                var outputPath = outputType.ToLower() == "pdf"
                    ? await GenerateSinglePdfAsync(session.TemplatePath, excelData, fileName)
                    : await GenerateSingleWordAsync(session.TemplatePath, excelData, fileName);
                
                if (outputPath != null)
                    session.OutputFiles.Add(Path.GetFileName(outputPath));
            }

            return ProcessingResult<List<string>>.Success(session.OutputFiles);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing merge");
            return ProcessingResult<List<string>>.Failure($"Error processing merge: {ex.Message}");
        }
    }

    private async Task<string?> GenerateSingleWordAsync(string templatePath, List<Dictionary<string, string>> data, string baseFileName)
    {
        try
        {
            var tempDir = Path.GetTempPath();
            var outputPath = Path.Combine(tempDir, $"{baseFileName}.docx");

            // Copy template as starting point
            File.Copy(templatePath, outputPath, true);

            using (var doc = WordprocessingDocument.Open(outputPath, true))
            {
                var mainPart = doc.MainDocumentPart;
                if (mainPart?.Document?.Body == null)
                    throw new InvalidOperationException("Invalid document structure");

                var body = mainPart.Document.Body;
                var originalContent = body.CloneNode(true);

                // Clear body content
                body.RemoveAllChildren();

                // Process each record
                for (int i = 0; i < data.Count; i++)
                {
                    var recordBody = (Body)originalContent.CloneNode(true);
                    
                    // Replace merge fields in this record's content
                    ReplaceMergeFields(recordBody, data[i]);
                    
                    // Add content to main body
                    foreach (var element in recordBody.Elements())
                    {
                        body.Append(element.CloneNode(true));
                    }

                    // Add page break between records (except after last)
                    if (i < data.Count - 1)
                    {
                        var pageBreak = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
                        body.Append(pageBreak);
                    }
                }

                mainPart.Document.Save();
            }

            return outputPath;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating Word document");
            return null;
        }
    }

    private async Task<string?> GenerateSinglePdfAsync(string templatePath, List<Dictionary<string, string>> data, string baseFileName)
    {
        try
        {
            // First generate Word document
            var wordPath = await GenerateSingleWordAsync(templatePath, data, baseFileName);
            if (wordPath == null) return null;

            // Convert Word to HTML
            var htmlContent = await ConvertWordToHtmlAsync(wordPath);
            if (string.IsNullOrEmpty(htmlContent)) return null;

            // Convert HTML to PDF using PuppeteerSharp
            var pdfPath = Path.Combine(Path.GetTempPath(), $"{baseFileName}.pdf");
            
            using var browser = await Puppeteer.LaunchAsync(new LaunchOptions 
            { 
                Headless = true,
                Args = new[] { "--no-sandbox", "--disable-setuid-sandbox" } // For Linux compatibility
            });
            
            using var page = await browser.NewPageAsync();
            await page.SetContentAsync(htmlContent);
            
            await page.PdfAsync(pdfPath, new PdfOptions
            {
                Format = PaperFormat.A4,
                DisplayHeaderFooter = false,
                MarginOptions = new MarginOptions
                {
                    Top = "20mm",
                    Right = "20mm",
                    Bottom = "20mm",
                    Left = "20mm"
                }
            });

            // Clean up Word file
            if (File.Exists(wordPath))
                File.Delete(wordPath);

            return pdfPath;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating PDF");
            return null;
        }
    }

    private async Task<string> ConvertWordToHtmlAsync(string wordPath)
    {
        try
        {
            using var doc = WordprocessingDocument.Open(wordPath, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            
            if (body == null) return string.Empty;

            var html = new StringBuilder();
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html><head>");
            html.AppendLine("<meta charset='UTF-8'>");
            html.AppendLine("<style>");
            html.AppendLine("body { font-family: 'Calibri', 'Arial', sans-serif; font-size: 12pt; line-height: 1.5; margin: 20px; }");
            html.AppendLine("h1 { font-size: 24pt; font-weight: bold; text-decoration: underline; margin-bottom: 20px; }");
            html.AppendLine("p { margin: 8px 0; }");
            html.AppendLine(".page-break { page-break-before: always; }");
            html.AppendLine("</style>");
            html.AppendLine("</head><body>");

            bool isFirstPage = true;
            foreach (var element in body.Elements())
            {
                if (element is Paragraph paragraph)
                {
                    var text = paragraph.InnerText;
                    
                    // Check for page breaks
                    var hasPageBreak = paragraph.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page);
                    if (hasPageBreak && !isFirstPage)
                    {
                        html.AppendLine("<div class='page-break'></div>");
                    }

                    // Check if this is a heading (contains "Invoice")
                    if (text.Trim().Equals("Invoice", StringComparison.OrdinalIgnoreCase))
                    {
                        html.AppendLine($"<h1>{text}</h1>");
                    }
                    else
                    {
                        // Check for bold formatting
                        var isBold = paragraph.Descendants<Bold>().Any();
                        if (isBold)
                        {
                            html.AppendLine($"<p><strong>{text}</strong></p>");
                        }
                        else
                        {
                            html.AppendLine($"<p>{text}</p>");
                        }
                    }
                    
                    isFirstPage = false;
                }
            }

            html.AppendLine("</body></html>");
            return html.ToString();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error converting Word to HTML");
            return string.Empty;
        }
    }

    private void ReplaceMergeFields(OpenXmlElement element, Dictionary<string, string> data)
    {
        foreach (var text in element.Descendants<Text>().ToList())
        {
            if (string.IsNullOrEmpty(text.Text)) continue;

            var updatedText = text.Text;
            foreach (var kvp in data)
            {
                var placeholder = $"{{{{{kvp.Key}}}}}";
                updatedText = updatedText.Replace(placeholder, kvp.Value);
            }
            text.Text = updatedText;
        }
    }

    private List<Dictionary<string, string>> LoadExcelData(string filePath)
    {
        var data = new List<Dictionary<string, string>>();
        
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets.FirstOrDefault();
        
        if (worksheet?.Dimension == null) return data;

        var colCount = worksheet.Dimension.Columns;
        var rowCount = worksheet.Dimension.Rows;

        // Get headers
        var headers = new List<string>();
        for (int col = 1; col <= colCount; col++)
        {
            var header = worksheet.Cells[1, col].Text;
            headers.Add(string.IsNullOrWhiteSpace(header) ? $"Column{col}" : header);
        }

        // Get data rows
        for (int row = 2; row <= rowCount; row++)
        {
            var rowData = new Dictionary<string, string>();
            for (int col = 1; col <= colCount; col++)
            {
                var value = worksheet.Cells[row, col].Text;
                rowData[headers[col - 1]] = value ?? "";
            }
            data.Add(rowData);
        }

        return data;
    }

    private string GetSafeFileName(Dictionary<string, string> record, List<string> headers, int index)
    {
        if (headers.Count > 0 && record.ContainsKey(headers[0]) && !string.IsNullOrWhiteSpace(record[headers[0]]))
        {
            var name = Regex.Replace(record[headers[0]], @"[<>:""/\\|?*]", "_");
            return $"{name}_{DateTime.Now:yyyyMMddHHmmss}";
        }
        return $"record_{index + 1}_{DateTime.Now:yyyyMMddHHmmss}";
    }

    public object GetSessionStatus(string sessionId)
    {
        var session = GetSession(sessionId);
        if (session == null)
            return new { hasTemplate = false, hasData = false };

        return new
        {
            hasTemplate = !string.IsNullOrEmpty(session.TemplatePath) && File.Exists(session.TemplatePath),
            hasData = !string.IsNullOrEmpty(session.DataPath) && File.Exists(session.DataPath),
            outputFiles = session.OutputFiles
        };
    }

    public async Task<ProcessingResult<DownloadFileInfo>> GetDownloadFileAsync(string sessionId, string filename)
    {
        try
        {
            var session = GetSession(sessionId);
            if (session == null || !session.OutputFiles.Contains(filename))
                return ProcessingResult<DownloadFileInfo>.Failure("File not found");

            var filePath = Path.Combine(Path.GetTempPath(), filename);
            if (!File.Exists(filePath))
                return ProcessingResult<DownloadFileInfo>.Failure("File not found");

            var extension = Path.GetExtension(filename).ToLowerInvariant();
            var contentType = extension switch
            {
                ".pdf" => "application/pdf",
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                _ => "application/octet-stream"
            };

            var fileInfo = new DownloadFileInfo
            {
                FilePath = filePath,
                FileName = filename,
                ContentType = contentType
            };

            return ProcessingResult<DownloadFileInfo>.Success(fileInfo);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting download file");
            return ProcessingResult<DownloadFileInfo>.Failure("Error accessing file");
        }
    }

    private MailMergeSession GetOrCreateSession(string sessionId)
    {
        if (!_sessions.ContainsKey(sessionId))
        {
            _sessions[sessionId] = new MailMergeSession();
            
            // Clean up old sessions if we have too many
            if (_sessions.Count > 50)
            {
                var oldestSession = _sessions.OrderBy(s => s.Value.CreatedAt).First();
                CleanupSession(oldestSession.Key);
                _sessions.Remove(oldestSession.Key);
            }
        }
        return _sessions[sessionId];
    }

    private MailMergeSession? GetSession(string sessionId)
    {
        return _sessions.ContainsKey(sessionId) ? _sessions[sessionId] : null;
    }

    private void CleanupSession(string sessionId)
    {
        if (!_sessions.ContainsKey(sessionId)) return;

        var session = _sessions[sessionId];
        
        // Clean up files
        if (!string.IsNullOrEmpty(session.TemplatePath) && File.Exists(session.TemplatePath))
            File.Delete(session.TemplatePath);
            
        if (!string.IsNullOrEmpty(session.DataPath) && File.Exists(session.DataPath))
            File.Delete(session.DataPath);

        foreach (var outputFile in session.OutputFiles)
        {
            var filePath = Path.Combine(Path.GetTempPath(), outputFile);
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }
}