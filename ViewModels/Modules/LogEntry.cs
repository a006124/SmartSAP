using System;
using System.Diagnostics; // Pour Process.Start
using System.IO;
using System.Windows.Input;

namespace SmartSAP.ViewModels.Modules
{
    public class LogEntry
    {
        public string Timestamp { get; private set; } = DateTime.Now.ToString("HH:mm:ss");
        public string Type { get; set; }
        public string Message { get; set; }
        public string? FilePath { get; set; }
        public string? FileName => !string.IsNullOrEmpty(FilePath) ? Path.GetFileName(FilePath) : null;
        public bool HasFile => !string.IsNullOrEmpty(FilePath);
        public string? LinkText { get; set; }
        public ICommand? LinkCommand { get; set; }
        public bool HasLink => !string.IsNullOrEmpty(LinkText) && LinkCommand != null;
        public ICommand? OpenFileCommand { get; }

        public LogEntry(string type, string message, string? filePath = null, string? linkText = null, ICommand? linkCommand = null)
        {
            Type = type;
            Message = message;
            FilePath = filePath;
            LinkText = linkText;
            LinkCommand = linkCommand;

            if (HasFile)
            {
                OpenFileCommand = new RelayCommand(_ =>
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(FilePath))
                            Process.Start(new ProcessStartInfo(FilePath) { UseShellExecute = true });
                    }
                    catch { /* Ignore errors on opening */ }
                });
            }
        }
    }
}
