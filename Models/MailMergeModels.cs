namespace MailMergeSaaS.Models;

public class MailMergeSession
{
    public string? TemplatePath { get; set; }
    public string? DataPath { get; set; }
    public List<string> Headers { get; set; } = new();
    public List<string> OutputFiles { get; set; } = new();
    public DateTime CreatedAt { get; set; } = DateTime.Now;
}

public class ProcessingResult<T>
{
    public bool Success { get; set; }
    public string? Error { get; set; }
    public T? Data { get; set; }

    public static ProcessingResult<T> Success(T data) => new() { Success = true, Data = data };
    public static ProcessingResult<T> Failure(string error) => new() { Success = false, Error = error };
}

public class DownloadFileInfo
{
    public string FilePath { get; set; } = string.Empty;
    public string FileName { get; set; } = string.Empty;
    public string ContentType { get; set; } = string.Empty;
}