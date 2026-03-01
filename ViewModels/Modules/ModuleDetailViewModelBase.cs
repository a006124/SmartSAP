using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Input;
using ClosedXML.Excel;
using SmartSAP.Models;
using System.Diagnostics;

namespace SmartSAP.ViewModels.Modules
{
    public abstract class ModuleDetailViewModelBase : ViewModelBase
    {
        protected readonly MainViewModel MainViewModel;

        public string ModuleTitle { get; protected set; }
        public ObservableCollection<WorkflowStep> Steps { get; protected set; }
        public ObservableCollection<LogEntry> Logs { get; protected set; }
        public ObservableCollection<ExcelColumnDefinition> ExcelColumns { get; protected set; }
        public ICommand GoBackCommand { get; protected set; }
        public ICommand RunWorkflowCommand { get; protected set; }
        public ICommand GenerateTemplateCommand { get; protected set; }

        protected ModuleDetailViewModelBase(MainViewModel mainViewModel, string title)
        {
            MainViewModel = mainViewModel;
            ModuleTitle = title;
            
            Logs = new ObservableCollection<LogEntry>();
            Steps = new ObservableCollection<WorkflowStep>();
            ExcelColumns = new ObservableCollection<ExcelColumnDefinition>();

            GoBackCommand = new RelayCommand(_ => MainViewModel.NavigateToLibrary());
            RunWorkflowCommand = new RelayCommand(async _ => await ExecuteWorkflowAsync());
            GenerateTemplateCommand = new RelayCommand(_ => GenerateExcelTemplate());
        }

        protected virtual void InitializeSteps()
        {
            // A surcharger dans les classes enfants pour définir les étapes spécifiques
        }

        protected virtual void InitializeExcelColumns()
        {
            // A surcharger dans les classes enfants pour définir les colonnes Excel
        }

        protected virtual void GenerateExcelTemplate()
        {
            if (ExcelColumns.Count == 0)
            {
                Logs.Add(new LogEntry("WARNING", "Aucun modèle Excel n'est défini pour ce module."));
                return;
            }

            try
            {
                // Note: Dans une application réelle on utiliserait un SaveFileDialog.
                // Ici on génère un nom par défaut pour la démonstration.
                string fileName = $"Template_{ModuleTitle.Replace(" ", "_")}.xlsx";
                string fullPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Modèle");
                    
                    for (int i = 0; i < ExcelColumns.Count; i++)
                    {
                        var colDef = ExcelColumns[i];
                        var cell = worksheet.Cell(1, i + 1);
                        
                        // Header
                        cell.Value = colDef.Header;
                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#3B82F6");
                        cell.Style.Font.FontColor = XLColor.White;

                        // Comment
                        if (!string.IsNullOrEmpty(colDef.Comment))
                        {
                            cell.GetComment().AddText(colDef.Comment);
                        }

                        // Sample Data
                        worksheet.Cell(2, i + 1).Value = colDef.SampleData;
                    }

                    worksheet.Columns().AdjustToContents();
                    workbook.SaveAs(fullPath);
                }

                Logs.Add(new LogEntry("SUCCESS", $"Modèle Excel généré avec succès sur le bureau : {fileName}"));

                // Ouverture automatique du fichier
                Process.Start(new ProcessStartInfo(fullPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Logs.Add(new LogEntry("ERROR", $"Erreur lors de la génération ou de l'ouverture du modèle : {ex.Message}"));
            }
        }

        protected void CompleteInitialization()
        {
            if (Steps.Count > 0)
            {
                foreach (var step in Steps) step.IsLast = false;
                Steps[Steps.Count - 1].IsLast = true;
            }
        }

        protected virtual async Task ExecuteWorkflowAsync()
        {
            Logs.Add(new LogEntry("INFO", $"Démarrage du workflow : {ModuleTitle}"));

            foreach (var step in Steps)
            {
                step.Status = "Processing";
                Logs.Add(new LogEntry("INFO", $"Exécution de : {step.Title}"));

                await Task.Delay(1500); // Simulation

                step.Status = "Completed";
                Logs.Add(new LogEntry("SUCCESS", $"{step.Title} terminé avec succès."));
            }

            Logs.Add(new LogEntry("INFO", "Workflow terminé."));
        }
    }

    public class WorkflowStep : ViewModelBase
    {
        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Icon { get; set; } = string.Empty;

        private string _status = "Ready";
        public string Status
        {
            get => _status;
            set => SetProperty(ref _status, value);
        }

        private bool _isLast;
        public bool IsLast
        {
            get => _isLast;
            set => SetProperty(ref _isLast, value);
        }

        public ICommand? ActionCommand { get; set; }
    }

    public class LogEntry
    {
        public string Timestamp { get; } = DateTime.Now.ToString("HH:mm:ss");
        public string Type { get; set; }
        public string Message { get; set; }

        public LogEntry(string type, string message)
        {
            Type = type;
            Message = message;
        }
    }
}
