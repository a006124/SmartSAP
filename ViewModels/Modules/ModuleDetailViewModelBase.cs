using ClosedXML.Excel;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Win32;
using SmartSAP.Models;
using SmartSAP.Services.SAP;
using SmartSAP.ViewModels;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics; // Pour Process.Start
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading; // Pour SynchronizationContext
using System.Threading.Tasks;
using System.Windows; // Pour Application
using System.Windows.Threading; // Pour Dispatcher
using System.Windows.Input; // Pour ICommand

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
        public ICommand GeneratePMPExcelCommand { get; protected set; }
        public ICommand ExportFixedWidthCommand { get; protected set; }
        public ICommand ClearLogsCommand { get; protected set; }
        public ICommand PickExcelFileCommand { get; protected set; }
        public ICommand CheckSAPConnectionCommand { get; protected set; }
        public ICommand ExecuteSAPTransactionCommand { get; protected set; }


        public ICommand GeneratePDFCommand { get; protected set; }
        //public ICommand GeneratePMPExcelCommand { get; protected set; }


        protected string? LastGeneratedExcelPath;
        protected string? LastExportedTextPath;
        protected string? LastGeneratedSAPExcelPath;
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
            GeneratePDFCommand = new RelayCommand(p => GeneratePDF(p as WorkflowStep));
            GeneratePMPExcelCommand = new RelayCommand(async p => await GeneratePMPExcel(p as WorkflowStep));

            ExportFixedWidthCommand = new RelayCommand(p =>
            {
                var step = p as WorkflowStep;
                if (step != null)
                {
                    ExportLastGeneratedToFixedWidth(step, step.NombreMini, step.OpenFile);
                }
                else
                {
                    // Gérer le cas où le paramètre n'est pas un WorkflowStep (par exemple, passer une valeur par défaut)
                    Logs.Add(new LogEntry("ERROR", "Le paramètre de la commande ExportFixedWidthCommand n'est pas un WorkflowStep valide."));
                    ExportLastGeneratedToFixedWidth(null, 0);
                }
            }); 
            ClearLogsCommand = new RelayCommand(_ => Logs.Clear());
            PickExcelFileCommand = new RelayCommand(_ => PickExcelFile());
            CheckSAPConnectionCommand = new RelayCommand(async p => await this.CheckSAPConnectionAsync());
            ExecuteSAPTransactionCommand = new RelayCommand(async p => await this.ExecuteSAPTransactionAsync(p as WorkflowStep));

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
                
                if (step != null) 
                { 
                    step.Status = "Terminé"; step.ResultState = "Success";
                    if (step.OpenFile)
                    {
                        // Ouverture automatique du fichier
                        Process.Start(new ProcessStartInfo(fullPath) { UseShellExecute = true });
                    }
                }

            }
            catch (Exception ex)
            {
                Logs.Add(new LogEntry("ERROR", $"Erreur lors de la génération ou de l'ouverture du modèle : {ex.Message}"));
                if (step != null) { step.Status = "Erreur"; step.ResultState = "Error"; }
            }
        }

        protected virtual void GeneratePDF(WorkflowStep? step = null)
        {
            if (step == null)
            {
                step = Steps.FirstOrDefault((WorkflowStep s) => s.ActionCommand == GeneratePDFCommand);
            }

            if (step != null) step.ResultState = "Processing";

            try
            {
                // Ouvrir une fenêtre de sélection de fichier
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Title = "Sélectionnez un fichier PDF";
                openFileDialog.Filter = "Fichiers PDF|*.pdf";

                if (openFileDialog.ShowDialog() == true)
                {
                    string inputPdf = openFileDialog.FileName;
                    string inputFileName=Path.GetFileNameWithoutExtension(inputPdf);
                    string outputDir = Path.GetDirectoryName(inputPdf);
                    int pagesPerFile = 20;

                    // Ouvrir le fichier PDF d'entrée
                    PdfReader reader = new PdfReader(inputPdf);

                    // Parcourir le PDF par segments de pagesPerFile
                    for (int i = 0; i < reader.NumberOfPages; i += pagesPerFile)
                    {
                        // Créer un nouveau document PDF
                        Document document = new Document();
                        string outputFilename = Path.Combine(outputDir, $"{inputFileName}_output_{(i / pagesPerFile + 1).ToString("D3")}.pdf");
                        PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(outputFilename, FileMode.Create));

                        document.Open();

                        // Ajouter les pages au nouveau fichier PDF
                        for (int j = i; j < Math.Min(i + pagesPerFile, reader.NumberOfPages); j++)
                        {
                            document.NewPage();
                            PdfImportedPage page = writer.GetImportedPage(reader, j + 1);
                            writer.DirectContent.AddTemplate(page, 0, 0);
                        }

                        document.Close();
                        writer.Close();

                        Console.WriteLine($"Créé: {outputFilename}");
                    }

                    Logs.Add(new LogEntry("SUCCESS", $"Fichiers PDF créés dans le dossier : ", outputDir));
                }

                if (step != null)
                {
                    step.Status = "Terminé"; step.ResultState = "Success";
                }

            }
            catch (Exception ex)
            {
                Logs.Add(new LogEntry("ERROR", $"Erreur lors de la création des fichiers PDF : {ex.Message}"));
                if (step != null) { step.Status = "Erreur"; step.ResultState = "Error"; }
            }
        }

        private void AddLog(LogEntry logEntry, Dispatcher dispatcher, SynchronizationContext uiSynchronizationContext)
        {
            try
            {
                System.Diagnostics.Trace.WriteLine($"[Trace] AddLog appelé pour : {logEntry.Type} - {logEntry.Message}. Dispatcher present? {dispatcher != null}");
                
                if (dispatcher != null)
                {
                    dispatcher.Invoke(() =>
                    {
                        System.Diagnostics.Trace.WriteLine($"[Trace] Ajout via Dispatcher: {logEntry.Type} - {logEntry.Message}");
                        Logs.Add(logEntry);
                    });
                }
                else if (uiSynchronizationContext != null)
                {
                    uiSynchronizationContext.Post(_ =>
                    {
                        System.Diagnostics.Trace.WriteLine($"[Trace] Ajout via SynchronizationContext: {logEntry.Type} - {logEntry.Message}");
                        Logs.Add(logEntry);
                    }, null);
                }
                else
                {
                    System.Diagnostics.Trace.WriteLine($"[Trace] Ajout direct (sans sync context): {logEntry.Type} - {logEntry.Message}");
                    Logs.Add(logEntry);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"[Trace] ERROR DANS AddLog : {ex.Message}");
            }
        }

        protected virtual async Task GeneratePMPExcel(WorkflowStep? step = null)
        {
            var uiSynchronizationContext = SynchronizationContext.Current;
            Dispatcher dispatcher = null;
            
            // Vérifiez si Application.Current est disponible (pour les applications WPF)
            if (Application.Current != null)
            {
                dispatcher = Application.Current.Dispatcher;
            }

            if (step == null)
            {
                step = Steps.FirstOrDefault((WorkflowStep s) => s.ActionCommand == GeneratePDFCommand);
            }

            if (step != null) step.ResultState = "Processing";

            AddLog(new LogEntry("INFO", "Préparation de la génération PMP..."), dispatcher, uiSynchronizationContext);
            await Task.Delay(10); // Rendre la main à l'UI pour afficher le message initial

            try
            {
                string docPath = Path.GetDirectoryName(LastGeneratedExcelPath) ?? AppDomain.CurrentDomain.BaseDirectory;

                string sFileName = "PMP_" + DateTime.Now.ToString("yyMMddHHmmss") + ".txt";
                int iLgLigne = 293;
                int iDebRechercheLigne = 270;
                int iLgFinRechercheLigne = 50;
                int endIndex;

                long rowNumber = 0;
                long rowTotalNumber;

                // Vérifier si le chemin du dossier n'est pas vide
                if (string.IsNullOrEmpty(docPath))
                {
                    throw new ArgumentException("Le chemin du dossier est invalide.");
                }

                // Obtenir tous les fichiers TXT dans le dossier spécifié, excluant ceux commençant par "PMP_"
                string[] fichiersCSV = Directory.GetFiles(docPath, "*.txt")
                                                .Where(f => !Path.GetFileName(f).StartsWith("PMP_"))
                                                .ToArray();

                // Vérifier si des fichiers TXT ont été trouvés
                if (fichiersCSV.Length == 0)
                {
                    throw new FileNotFoundException("Aucun fichier TXT trouvé dans le dossier spécifié.");
                }
                else
                {
                    rowTotalNumber = fichiersCSV.Length;
                }

                // Utiliser un StreamWriter pour écrire dans le fichier de sortie de manière asynchrone
                using (StreamWriter writer = new StreamWriter(Path.Combine(docPath, sFileName), false)) // Le paramètre 'false' écrase le fichier s'il existe
                {
                    StringBuilder accumulatedLine = new StringBuilder();
                    Regex regex = new Regex(@"[^\w\s/]"); // Expression régulière pour remplacer les caractères non alphanumériques sauf '/'

                    foreach (string fichier in fichiersCSV)
                    {
                        rowNumber += 1;

                        // Lire toutes les lignes du fichier
                        string[] lignesFichier = File.ReadAllLines(fichier);

                        // Parcourir chaque ligne du fichier
                        foreach (string ligne in lignesFichier)
                        {
                            // Remplacer les caractères non alphanumériques
                            string cleanedLine = regex.Replace(ligne, " ");
                            // Ajouter la ligne nettoyée à la ligne accumulée
                            accumulatedLine.Append(cleanedLine);
                        }
                    }

                    // Écrire des lignes de 293 caractères de longueur
                    // On soustrait 10 secondes pour forcer un premier affichage immédiat
                    DateTime lastLogTime = DateTime.Now.AddSeconds(-10); 
                    try
                    {
                        while (accumulatedLine.Length > iLgLigne)
                        {
                            Match match = null;
                            string pattern = ""; // Initialiser pattern
                            endIndex = 0;

                            // Calculer la longueur de la sous-chaîne à rechercher, en évitant les erreurs d'index
                            int searchLength = Math.Min(iLgFinRechercheLigne, accumulatedLine.Length - iDebRechercheLigne);
                            if (searchLength > 0)
                            {
                                for (int i = 1; i <= 99; i++)
                                {
                                    pattern = " " + i.ToString("D2");
                                    match = Regex.Match(accumulatedLine.ToString(iDebRechercheLigne, searchLength), pattern);
                                    if (match.Success)
                                    {
                                        endIndex = iDebRechercheLigne + match.Index + pattern.Length;
                                        break; // Sortir de la boucle si un match est trouvé
                                    }
                                }
                            }

                            // Si aucun match n'est trouvé ou si l'endIndex est invalide, prendre la longueur maximale de la ligne
                            if (endIndex == 0 || endIndex > accumulatedLine.Length)
                            {
                                endIndex = Math.Min(iLgLigne, accumulatedLine.Length);
                            }

                            string lineToWrite = accumulatedLine.ToString(0, endIndex);
                            accumulatedLine.Remove(0, endIndex);

                            // Ajouter des espaces avant le nombre de deux chiffres pour atteindre la longueur souhaitée
                            if (lineToWrite.Length < iLgLigne)
                            {
                                int paddingLength = iLgLigne - lineToWrite.Length;
                                int paddingIndex = lineToWrite.LastIndexOf(" ");
                                if (paddingIndex > 0)
                                {
                                    lineToWrite = lineToWrite.Insert(paddingIndex, new string(' ', paddingLength));
                                }
                            }
                            else if (lineToWrite.Length >= iLgLigne)
                            {
                                // Elle tronque la ligne et ajoute le dernier 'pattern' trouvé ou généré.
                                lineToWrite = lineToWrite.Substring(0, iLgLigne - 3) + pattern;
                            }

                            const int MaxRetries = 5; // Nombre maximal de tentatives
                            const int DelayMs = 100;  // Délai entre les tentatives en millisecondes
                            int retry = 0;
                            for ( retry = 0; retry < MaxRetries; retry++)
                            {
                                try
                                {
                                    await Task.Delay(DelayMs); // Attendre de manière asynchrone pour ne pas bloquer l'UI
                                    await writer.WriteLineAsync(lineToWrite); // Utiliser await pour attendre la fin de l'écriture
                                    break; // Si l'écriture réussit, sortir de la boucle de retry
                                }
                                catch (IOException ex)
                                {
                                    // Le HResult -2147024864 correspond à "The process cannot access the file because it is being used by another process."
                                    // Vous pouvez aussi vérifier d'autres codes d'erreur si nécessaire.
                                    if (ex.HResult == -2147024864 && retry < MaxRetries - 1)
                                    {
                                        // Loguer l'échec temporaire si vous voulez
                                        AddLog(new LogEntry("WARNING", $"Tentative {retry + 1}/{MaxRetries} d'écriture échouée (fichier en cours d'utilisation). Réessai dans {DelayMs}ms. Erreur: {ex.Message}"), dispatcher, uiSynchronizationContext);
                                        await Task.Delay(DelayMs); // Attendre de manière asynchrone
                                    }
                                    else
                                    {
                                        // Si c'est la dernière tentative ou un autre type d'IOException, relancer l'exception
                                        throw;
                                    }
                                } 
                            }

                            // Log "Traitement en cours ..." toutes les 10 secondes
                            if ((DateTime.Now - lastLogTime).TotalSeconds >= 10)
                            {
                                AddLog(new LogEntry("INFO", "Traitement en cours ..."), dispatcher, uiSynchronizationContext);
                                lastLogTime = DateTime.Now; // Réinitialiser le temps du dernier log
                                
                                // Rendre la main au thread UI pour que l'affichage puisse se mettre à jour
                                await Task.Delay(10);
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        AddLog(new LogEntry("ERROR", $"Erreur lors de la création du fichier PMP Excel : {ex.Message}"), dispatcher, uiSynchronizationContext);
                        if (step != null) { step.Status = "Erreur"; step.ResultState = "Error"; }
                    }
                }

                AddLog(new LogEntry("SUCCESS", $"Fichiers PMP créés dans le dossier : " + docPath), dispatcher, uiSynchronizationContext);

                if (step != null)
                {
                    step.Status = "Terminé"; step.ResultState = "Success";
                }

            }
            catch (Exception ex)
            {
                AddLog(new LogEntry("ERROR", $"Erreur lors de la création du fichier PMP Excel : {ex.Message}"), dispatcher, uiSynchronizationContext);
                if (step != null) { step.Status = "Erreur"; step.ResultState = "Error"; }
            }
        }


        protected virtual void ExportLastGeneratedToFixedWidth(WorkflowStep? step = null, int nombreMini = 0, bool OpenFile = true)
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

                    if (rows.Count() < nombreMini)
                    {
                        Logs.Add(new LogEntry("WARNING", $"Le nombre de ligne à traiter doit être supérieur ou égal à {nombreMini}."));
                        if (step != null) { step.Status = "Données insuffisantes"; step.ResultState = "Error"; }
                        return;
                    }

                    using (var writer = new StreamWriter(exportPath, false, System.Text.Encoding.UTF8))

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

                                if (!colDef.ForcerVide)
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
                                        var rules = colDef.RègleDeGestion.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(r => r.Trim());
                                        foreach (var règle in rules)
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
                        if (step != null) 
                        { 
                            step.Status = "Terminé"; 
                            step.ResultState = "Success";
                            if (OpenFile)
                            {
                                // Ouverture automatique de l'export
                                Process.Start(new ProcessStartInfo(exportPath) { UseShellExecute = true });
                            }
                        }

                        LastExportedTextPath = exportPath;

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
                Logs.Add(new LogEntry("ERROR", "Veuillez d'abord générer l'export format SAP."));
                if (step != null) { step.Status = "Manquant"; step.ResultState = "Error"; }
                return;
            }

            Logs.Add(new LogEntry("INFO", "Exécution de la transaction SAP..."));
            // L'implémentation spécifique par module se fera par surcharge dans les ViewModels enfants.
            await Task.CompletedTask; 
        }

        protected virtual void PickExcelFile()
        {
            var openFileDialog = new OpenFileDialog
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
