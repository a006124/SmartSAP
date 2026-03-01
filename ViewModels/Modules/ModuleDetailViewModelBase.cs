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
        public ICommand ExportFixedWidthCommand { get; protected set; }
        public ICommand ClearLogsCommand { get; protected set; }
        public ICommand PickExcelFileCommand { get; protected set; }

        protected string? LastGeneratedExcelPath;

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
            ExportFixedWidthCommand = new RelayCommand(_ => ExportLastGeneratedToFixedWidth());
            ClearLogsCommand = new RelayCommand(_ => Logs.Clear());
            PickExcelFileCommand = new RelayCommand(_ => PickExcelFile());
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
            var step = Steps.FirstOrDefault(s => s.ActionCommand == GenerateTemplateCommand);
            if (step != null) step.ResultState = "Processing";

            if (ExcelColumns.Count == 0)
            {
                Logs.Add(new LogEntry("WARNING", "Aucun modèle Excel n'est défini pour ce module."));
                if (step != null) { step.Status = "Warning"; step.ResultState = "Error"; }
                return;
            }

            try
            {
                // Note: Dans une application réelle on utiliserait un SaveFileDialog.
                // Ici on génère un nom par défaut pour la démonstration.
                string dateExecution = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string fileName = $"{dateExecution}_{ModuleTitle.Replace(" ", "_")}.xlsx";
                string fullPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);
                string sheetName = "Data";
                
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(sheetName);
                    
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

                LastGeneratedExcelPath = fullPath;
                Logs.Add(new LogEntry("SUCCESS", $"Modèle Excel généré avec succès sur le bureau : ", fullPath));
                
                if (step != null) { step.Status = "Terminé"; step.ResultState = "Success"; }

                // Ouverture automatique du fichier
                Process.Start(new ProcessStartInfo(fullPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Logs.Add(new LogEntry("ERROR", $"Erreur lors de la génération ou de l'ouverture du modèle : {ex.Message}"));
                if (step != null) { step.Status = "Erreur"; step.ResultState = "Error"; }
            }
        }

        protected virtual void ExportLastGeneratedToFixedWidth()
        {
            var step = Steps.FirstOrDefault(s => s.ActionCommand == ExportFixedWidthCommand);
            if (step != null) step.ResultState = "Processing";

            if (string.IsNullOrEmpty(LastGeneratedExcelPath) || !File.Exists(LastGeneratedExcelPath))
            {
                Logs.Add(new LogEntry("WARNING", "Aucun fichier Excel récent n'a été trouvé pour l'export. veuillez d'abord générer ou modifier le fichier."));
                if (step != null) { step.Status = "Absent"; step.ResultState = "Error"; }
                return;
            }

            try
            {
                string exportPath = Path.ChangeExtension(LastGeneratedExcelPath, ".txt");
                
                using (var workbook = new XLWorkbook(LastGeneratedExcelPath))
                {
                    var worksheet = workbook.Worksheets.First();
                    var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Ignorer l'en-tête
                    int errorCount = 0;
                    int rowIdx = 1;

                    using (var writer = new StreamWriter(exportPath))
                    {
                        foreach (var row in rows)
                        {
                            rowIdx++;
                            string line = "";
                            bool rowValid = true;

                            for (int i = 0; i < ExcelColumns.Count; i++)
                            {
                                var colDef = ExcelColumns[i];
                                int width = colDef.FixedWidth;
                                string rawValue = row.Cell(i + 1).Value.ToString();
                                
                                // Cleansing
                                string processedValue = rawValue?.Trim() ?? "";
                                if (colDef.ForceUpperCase)
                                    processedValue = processedValue.ToUpper();

                                // Validation
                                if (colDef.AllowedValues != null && colDef.AllowedValues.Length > 0)
                                {
                                    bool match = false;
                                    foreach (var allowed in colDef.AllowedValues)
                                    {
                                        if (processedValue == allowed.ToUpper())
                                        {
                                            match = true;
                                            break;
                                        }
                                    }

                                    if (!match)
                                    {
                                        Logs.Add(new LogEntry("ERROR", $"Ligne {rowIdx} : Valeur '{processedValue}' non autorisée pour '{colDef.Header}'."));
                                        rowValid = false;
                                        errorCount++;
                                    }
                                }
                                
                                if (width > 0)
                                {
                                    // Tronquer ou Padder à droite
                                    if (processedValue.Length > width)
                                        processedValue = processedValue.Substring(0, width);
                                    else
                                        processedValue = processedValue.PadRight(width);
                                    
                                    line += processedValue;
                                }
                                else
                                {
                                    line += processedValue + " ";
                                }
                            }

                            if (rowValid)
                                writer.WriteLine(line);
                        }
                    }

                    if (errorCount > 0)
                    {
                        Logs.Add(new LogEntry("WARNING", $"Export terminé avec {errorCount} erreur(s). Les lignes erronées ont été ignorées."));
                        if (step != null) { step.Status = "Erreurs"; step.ResultState = "Error"; }

                        // Ouverture automatique du fichier Excel pour correction
                        if (!string.IsNullOrEmpty(LastGeneratedExcelPath) && File.Exists(LastGeneratedExcelPath))
                        {
                            Logs.Add(new LogEntry("INFO", "Ouverture du fichier Excel pour correction..."));
                            Process.Start(new ProcessStartInfo(LastGeneratedExcelPath) { UseShellExecute = true });
                        }
                    }
                    else
                    {
                        Logs.Add(new LogEntry("SUCCESS", $"Export format SAP (taille fixe) généré avec succès : ", exportPath));
                        if (step != null) { step.Status = "Terminé"; step.ResultState = "Success"; }

                        // Ouverture automatique de l'export
                        Process.Start(new ProcessStartInfo(exportPath) { UseShellExecute = true });
                    }
                }
            }
            catch (Exception ex)
            {
                Logs.Add(new LogEntry("ERROR", $"Erreur lors de l'export fixe : {ex.Message}"));
                if (step != null) { step.Status = "Erreur"; step.ResultState = "Error"; }
            }
        }

        protected virtual void PickExcelFile()
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Title = "Sélectionner le fichier de données"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                LastGeneratedExcelPath = openFileDialog.FileName;
                Logs.Add(new LogEntry("INFO", $"Fichier source sélectionné : {Path.GetFileName(LastGeneratedExcelPath)}"));
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
                step.ResultState = "Processing";
                step.Status = "En cours";
                Logs.Add(new LogEntry("INFO", $"Exécution de : {step.Title}"));

                await Task.Delay(1500); // Simulation

                step.ResultState = "Success";
                step.Status = "Terminé";
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

        private string _resultState = "Normal";
        public string ResultState
        {
            get => _resultState;
            set => SetProperty(ref _resultState, value);
        }

        public string? LinkText { get; set; }
        public ICommand? LinkCommand { get; set; }

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
        public string Timestamp { get; private set; } = DateTime.Now.ToString("HH:mm:ss");
        public string Type { get; set; }
        public string Message { get; set; }
        public string? FilePath { get; set; }
        public string? FileName => !string.IsNullOrEmpty(FilePath) ? Path.GetFileName(FilePath) : null;
        public bool HasFile => !string.IsNullOrEmpty(FilePath);
        public ICommand? OpenFileCommand { get; }

        public LogEntry(string type, string message, string? filePath = null)
        {
            Type = type;
            Message = message;
            FilePath = filePath;
            
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
