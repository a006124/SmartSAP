using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using NPOI.SS.Formula.Functions;
using SmartSAP.Models;
using SmartSAP.Services.SAP;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Input;

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
                string dateExecution = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string moduleStepPart = !string.IsNullOrEmpty(step?.ModuleStep) ? $"{step.ModuleStep}" : "";
                string fileName = $"{dateExecution}_{ModuleTitle.Replace(" ", "_")}_{moduleStepPart}.xlsx";
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
                        cell.Value = colDef.Entete;
                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#3B82F6");
                        cell.Style.Font.FontColor = XLColor.White;

                        // Comment
                        if (!string.IsNullOrEmpty(colDef.Commentaires))
                        {
                            cell.GetComment().AddText(colDef.Commentaires);
                        }

                        // Sample Data
                        worksheet.Cell(2, i + 1).Value = colDef.Exemple;
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

            InitializeExcelColumns(step);

            if (ExcelColumns.Count == 0)
            {
                Logs.Add(new LogEntry("WARNING", "Aucun modèle Excel n'est défini pour cette étape."));
                if (step != null) { step.Status = "Warning"; step.ResultState = "Error"; }
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

                    if (rows.Count() < 2)
                    {
                        Logs.Add(new LogEntry("WARNING", "Le nombre de ligne à traiter doit être supérieur ou égal à 2."));
                        if (step != null) { step.Status = "Données insuffisantes"; step.ResultState = "Error"; }
                        return;
                    }

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
                                int width = colDef.LongueurMaxi;
                                string rawValue = row.Cell(i + 1).Value.ToString();
                                string processedValue = string.Empty;

                                if (!colDef.ForcerVide){
                                {
                                    // Cleaning
                                    if (rawValue != null) {
                                        processedValue = rawValue.Replace("\r", " ").Replace("\n", " "); // Remplacer les retours à la ligne par des espaces pour éviter de casser le format fixe                                            
                                        processedValue = processedValue?.Trim() ?? ""; // Suppression des espaces superflus
                                    }
                                    else {
                                        rawValue = string.Empty;
                                    }

                                    // Majuscules si nécessaire
                                    if (colDef.ForcerMajuscule)
                                        processedValue = processedValue.ToUpper();

                                    // Valeurs autorisées ?
                                    if (colDef.ValeursAutorisées != null && colDef.ValeursAutorisées.Length > 0)
                                    {
                                        bool match = false;
                                        foreach (var allowed in colDef.ValeursAutorisées)
                                        {
                                            if (processedValue == allowed.ToUpper())
                                            {
                                                match = true;
                                                break;
                                            }
                                        }

                                       if (!match)
                                       {
                                            string allowedStr = string.Join(", ", colDef.ValeursAutorisées);
                                            Logs.Add(new LogEntry("ERROR", $"Ligne {rowIdx} : Valeur '{processedValue}' non autorisée pour '{colDef.Entete}'. Valeur attendue : {allowedStr}"));
                                            rowValid = false;
                                            errorCount++;
                                       }
                                    }

                                    // Règles de gestion
                                    if (colDef.RègleDeGestion != null && colDef.RègleDeGestion.Length > 0)
                                    {
                                        foreach (var règle in colDef.RègleDeGestion)
                                        {
                                            switch (règle) // Mnn.n.A : Fichier utilisé dans le module Mnn Etape n colonne A
                                            {
                                                case "M04.2.W":
                                                case "M05.1.2.C":
                                                case "M05.3.W":
                                                case "M05.3.AK": // Doit être numérique
                                                    bool match = false;
                                                    if (string.IsNullOrEmpty(processedValue))
                                                        match = true; // Autoriser les champs vides
                                                    else
                                                        match = Regex.IsMatch(processedValue, @"^\d+$");
                                                    if (!match)
                                                    {
                                                        Logs.Add(new LogEntry("ERROR", $"Ligne {rowIdx} : Valeur '{processedValue}' non autorisée pour '{colDef.Entete}'. Valeur attendue : uniquement des chiffres"));
                                                        rowValid = false;
                                                        errorCount++;
                                                    }
                                                    break;
                                                case "M01.2.G":
                                                case "M04.2.J":
                                                case "M05.3.J": // Doit être au format 9999
                                                    if (string.IsNullOrEmpty(processedValue))
                                                        match = true; // Autoriser les champs vides
                                                    else
                                                        match = Regex.IsMatch(processedValue, @"^\d{4}$");
                                                    if (!match)
                                                    {
                                                        Logs.Add(new LogEntry("ERROR", $"Ligne {rowIdx} : Valeur '{processedValue}' non autorisée pour '{colDef.Entete}'. Valeur attendue : uniquement un nombre au format 9999"));
                                                        rowValid = false;
                                                        errorCount++;
                                                    }
                                                    break;
                                                    case "M04.2.Y":
                                                    case "M04.2.Z":
                                                    case "M04.2.AA":
                                                    case "M04.2.AW":
                                                    case "M05.3.Y":
                                                    case "M05.3.Z":
                                                    case "M05.3.AA":
                                                    case "M05.3.AW": // Doit être au format JJMMAAAA
                                                        if (string.IsNullOrEmpty(processedValue))
                                                            match = true; // Autoriser les champs vides
                                                        else
                                                            {
                                                                string pattern = @"^(0[1-9]|[12][0-9]|3[01])(0[1-9]|1[0-2])\d{4}$";
                                                                match = Regex.IsMatch(processedValue,pattern);
                                                            }
                                                        if (!match)
                                                        {
                                                            Logs.Add(new LogEntry("ERROR", $"Ligne {rowIdx} : Valeur '{processedValue}' non autorisée pour '{colDef.Entete}'. Valeur attendue : une date au format JJMMAAA"));
                                                            rowValid = false;
                                                            errorCount++;
                                                        }
                                                        break;
                                                case "M04.2.AD":
                                                case "M05.3.AD": // Doit être au format code MABEC
                                                    if (string.IsNullOrEmpty(processedValue))
                                                        match = true; // Autoriser les champs vides
                                                    else
                                                    {
                                                        string pattern = @"^.{10}$";
                                                        match = Regex.IsMatch(processedValue, pattern);
                                                    }
                                                    if (!match)
                                                    {
                                                        Logs.Add(new LogEntry("ERROR", $"Ligne {rowIdx} : Valeur '{processedValue}' non autorisée pour '{colDef.Entete}'. Valeur attendue : une chaine de 10 caractères"));
                                                        rowValid = false;
                                                        errorCount++;
                                                    }
                                                    break;
                                                case "M04.2.AP":
                                                case "M05.3.AP": // 6 caractères numériques ou ZZZBDN
                                                    if (string.IsNullOrEmpty(processedValue))
                                                        match = true; // Autoriser les champs vides
                                                    else
                                                    {
                                                        string pattern = @"^(\d{6}|ZZZBDN)$";
                                                        match = Regex.IsMatch(processedValue, pattern);
                                                    }
                                                    if (!match)
                                                    {
                                                        Logs.Add(new LogEntry("ERROR", $"Ligne {rowIdx} : Valeur '{processedValue}' non autorisée pour '{colDef.Entete}'. Valeur attendue : une chaine de 6 chiffres ou ZZZBFN"));
                                                        rowValid = false;
                                                        errorCount++;
                                                    }
                                                    break;

                                            }
                                        }
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

        public virtual void HandleDroppedFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;

            string extension = Path.GetExtension(filePath).ToLowerInvariant();

            if (extension == ".xlsx" || extension == ".xls")
            {
                HandleDroppedExcelFile(filePath);
            }
            else if (extension == ".txt" || extension == ".csv")
            {
                HandleDroppedTexteFile(filePath);
            }
            else
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    Logs.Add(new LogEntry(
                        "WARNING", 
                        $"Format de fichier non pris en charge déposé : {Path.GetFileName(filePath)}"
                    ));
                });
            }
        }

        public virtual void HandleDroppedExcelFile(string filePath)
        {
            LastGeneratedExcelPath = filePath;
            
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                Logs.Add(new LogEntry(
                    "INFO", 
                    "Fichier Excel déposé manuellement : ", 
                    filePath
                ));
            });
        }

        public virtual void HandleDroppedTexteFile(string filePath)
        {
            LastExportedTextPath = filePath;
            
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                Logs.Add(new LogEntry(
                    "INFO", 
                    "Fichier texte/CSV déposé manuellement : ", 
                    filePath
                ));
            });
        }
    }
}
