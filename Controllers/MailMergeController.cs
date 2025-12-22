using Microsoft.AspNetCore.Mvc;
using MailMergeSaaS.Services;
using MailMergeSaaS.Models;

namespace MailMergeSaaS.Controllers;

public class MailMergeController : Controller
{
    private readonly MailMergeService _mailMergeService;
    private readonly ILogger<MailMergeController> _logger;

    public MailMergeController(MailMergeService mailMergeService, ILogger<MailMergeController> logger)
    {
        _mailMergeService = mailMergeService;
        _logger = logger;
    }

    public IActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public async Task<IActionResult> UploadTemplate(IFormFile template)
    {
        try
        {
            if (template == null || template.Length == 0)
                return Json(new { success = false, error = "No template file provided" });

            if (!Path.GetExtension(template.FileName).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                return Json(new { success = false, error = "Only .docx files are supported" });

            var sessionId = GetOrCreateSessionId();
            var result = await _mailMergeService.UploadTemplateAsync(sessionId, template);

            return Json(new { success = result.Success, error = result.Error, templatePath = result.Data });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error uploading template");
            return Json(new { success = false, error = "Server error occurred" });
        }
    }

    [HttpPost]
    public async Task<IActionResult> UploadData(IFormFile data)
    {
        try
        {
            if (data == null || data.Length == 0)
                return Json(new { success = false, error = "No data file provided" });

            if (!Path.GetExtension(data.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                return Json(new { success = false, error = "Only .xlsx files are supported" });

            var sessionId = GetOrCreateSessionId();
            var result = await _mailMergeService.UploadDataAsync(sessionId, data);

            return Json(new { 
                success = result.Success, 
                error = result.Error, 
                preview = result.Data,
                dataPath = result.Success ? "uploaded" : null 
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error uploading data");
            return Json(new { success = false, error = "Server error occurred" });
        }
    }

    [HttpPost]
    public async Task<IActionResult> ProcessMerge([FromBody] ProcessMergeRequest request)
    {
        try
        {
            var sessionId = GetOrCreateSessionId();
            var result = await _mailMergeService.ProcessMergeAsync(sessionId, request.OutputType, request.MultipleFiles);

            if (result.Success)
            {
                return Json(new { 
                    success = true, 
                    files = result.Data,
                    message = "Processing completed successfully"
                });
            }

            return Json(new { success = false, error = result.Error });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing merge");
            return Json(new { success = false, error = "Server error occurred" });
        }
    }

    [HttpGet]
    public IActionResult CheckStatus()
    {
        try
        {
            var sessionId = GetOrCreateSessionId();
            var status = _mailMergeService.GetSessionStatus(sessionId);
            return Json(status);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error checking status");
            return Json(new { hasTemplate = false, hasData = false, error = "Server error occurred" });
        }
    }

    [HttpGet]
    public async Task<IActionResult> Download(string filename)
    {
        try
        {
            var sessionId = GetOrCreateSessionId();
            var result = await _mailMergeService.GetDownloadFileAsync(sessionId, filename);

            if (!result.Success || result.Data == null)
                return NotFound("File not found");

            var fileInfo = result.Data;
            return File(System.IO.File.ReadAllBytes(fileInfo.FilePath), fileInfo.ContentType, fileInfo.FileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error downloading file: {Filename}", filename);
            return NotFound("File not found");
        }
    }

    private string GetOrCreateSessionId()
    {
        var sessionId = HttpContext.Session.GetString("SessionId");
        if (string.IsNullOrEmpty(sessionId))
        {
            sessionId = Guid.NewGuid().ToString();
            HttpContext.Session.SetString("SessionId", sessionId);
        }
        return sessionId;
    }

    public class ProcessMergeRequest
    {
        public string OutputType { get; set; } = string.Empty;
        public bool MultipleFiles { get; set; }
    }
}