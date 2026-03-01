using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Input;
using ClosedXML.Excel;
using SmartSAP.Models;
using System.Diagnostics;
using System.Linq;
using SmartSAP.Services.SAP;

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
        public ICommand GenerateTemplateCommand { get; protected set; }
        public ICommand ExportFixedWidthCommand { get; protected set; }
        public ICommand ClearLogsCommand { get; protected set; }
        public ICommand PickExcelFileCommand { get; protected set; }
        public ICommand CheckSAPConnectionCommand { get; protected set; }
        public ICommand ExecuteSAPTransactionCommand { get; protected set; }

        protected string? LastGeneratedExcelPath;
        protected string? LastExportedTextPath;
        protected readonly SAPManager SAPManager;

        protected ModuleDetailViewModelBase(MainViewModel mainViewModel, string title)
        {
            MainViewModel = mainViewModel;
            ModuleTitle = title;
            
            Logs = new ObservableCollection<LogEntry>();
            Steps = new ObservableCollection<WorkflowStep>();
            ExcelColumns = new ObservableCollection<ExcelColumnDefinition>();

            GoBackCommand = new RelayCommand(_ => MainViewModel.NavigateToLibrary());
            GenerateTemplateCommand = new RelayCommand(p => GenerateExcelTemplate(p as WorkflowStep));
            ExportFixedWidthCommand = new RelayCommand(p => ExportLastGeneratedToFixedWidth(p as WorkflowStep));
            ClearLogsCommand = new RelayCommand(_ => Logs.Clear());
            PickExcelFileCommand = new RelayCommand(_ => PickExcelFile());
            CheckSAPConnectionCommand = new RelayCommand(async _ => await CheckSAPConnectionAsync());
            ExecuteSAPTransactionCommand = new RelayCommand(async p => await ExecuteSAPTransactionAsync(p as WorkflowStep));

            SAPManager = new SAPManager();

            InitializeSteps();
            CompleteInitialization();
        }

        protected virtual void InitializeSteps()
        {
            // A surcharger dans les classes enfants pour définir les étapes spécifiques
        }

        protected virtual void InitializeExcelColumns(WorkflowStep? step = null)
        {
            // A surcharger dans les classes enfants pour définir les colonnes Excel
        }

        protected virtual void GenerateExcelTemplate(WorkflowStep? step = null)
        {
            if (step == null)
            {
                step = Steps.FirstOrDefault((WorkflowStep s) => s.ActionCommand == GenerateTemplateCommand);
            }

            if (step != null) step.ResultState = "Processing";

            InitializeExcelColumns(step);

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
                string moduleStepPart = !string.IsNullOrEmpty(step?.ModuleStep) ? $"_{step.ModuleStep}" : "";
                string fileName = $"{dateExecution}_{ModuleTitle.Replace(" ", "_")}_{moduleStepPart}_{dateExecution}.xlsx";
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

        protected virtual void ExportLastGeneratedToFixedWidth(WorkflowStep? step = null)
        {
            if (step == null)
            {
                step = Steps.FirstOrDefault(s => s.ActionCommand == ExportFixedWidthCommand);
            }
            
            if (step != null) step.ResultState = "Processing";

            if (string.IsNullOrEmpty(LastGeneratedExcelPath) || !File.Exists(LastGeneratedExcelPath))
            {
                Logs.Add(new LogEntry("WARNING", "Aucun fichier Excel récent n'a été trouvé pour l'export. Créer un modèle Excel ou ", null, "modifier le fichier.", PickExcelFileCommand));
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
                        
                        LastExportedTextPath = exportPath;

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

        protected virtual async Task CheckSAPConnectionAsync()
        {
            var step = Steps.FirstOrDefault((WorkflowStep s) => s.ActionCommand == CheckSAPConnectionCommand);
            if (step != null) { step.ResultState = "Processing"; step.Status = "Vérification..."; }

            Logs.Add(new LogEntry("INFO", "Vérification de la connexion SAP en cours..."));
            
            await Task.Run(() =>
            {
                var result = SAPManager.IsConnectedToSAP();
                
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    if (result.IsSuccess)
                    {
                        Logs.Add(new LogEntry("SUCCESS", $"✓ Connexion SAP OK. Instance : {result.InstanceInfo}"));
                        if (step != null) { step.Status = "Connecté"; step.ResultState = "Success"; }
                        
                        MainViewModel.IsSAPConnected = true;
                        MainViewModel.SAPInstanceInfo = $"Instance : {result.InstanceInfo}";
                    }
                    else
                    {
                        Logs.Add(new LogEntry("ERROR", result.ErrorMessage));
                        if (step != null) { step.Status = "Erreur"; step.ResultState = "Error"; }
                        
                        MainViewModel.IsSAPConnected = false;
                        MainViewModel.SAPInstanceInfo = "Non connecté";
                    }
                });
            });
        }

        protected virtual async Task ExecuteSAPTransactionAsync(WorkflowStep? step = null)
        {
            if (step == null)
            {
                step = Steps.FirstOrDefault(s => s.ActionCommand == ExecuteSAPTransactionCommand);
            }
            
            if (step != null) step.ResultState = "Processing";

            if (string.IsNullOrEmpty(LastExportedTextPath) || !File.Exists(LastExportedTextPath))
            {
                Logs.Add(new LogEntry("ERROR", "Veuillez d'abord générer l'export format SAP (Étape 2)."));
                if (step != null) { step.Status = "Manquant"; step.ResultState = "Error"; }
                return;
            }

            Logs.Add(new LogEntry("INFO", "Exécution de la transaction SAP..."));
            // L'implémentation spécifique par module se fera par surcharge dans les ViewModels enfants.
            await Task.CompletedTask; 
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

                try
                {
                    Process.Start(new ProcessStartInfo(LastGeneratedExcelPath) { UseShellExecute = true });
                }
                catch { /* Ignorer les erreurs d'ouverture */ }
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
    }
}
