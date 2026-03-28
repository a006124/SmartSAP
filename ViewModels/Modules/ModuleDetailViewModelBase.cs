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
using System.Windows.Input; // Pour ICommand
using System.Windows.Threading; // Pour Dispatcher

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

        protected void AddLog(LogEntry logEntry, Dispatcher dispatcher = null, SynchronizationContext uiSynchronizationContext = null)
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

            if (Application.Current != null)
            {
                dispatcher = Application.Current.Dispatcher;
            }

            if (step == null)
            {
                step = Steps.FirstOrDefault((WorkflowStep s) => s.ActionCommand == GeneratePDFCommand);
            }

            if (step != null) step.ResultState = "Processing";

            try
            {
                string docPath = Path.GetDirectoryName(LastGeneratedExcelPath) ?? AppDomain.CurrentDomain.BaseDirectory;
                string sFileName = "PMP_" + DateTime.Now.ToString("yyMMddHHmmss") + ".txt";

                bool txtSuccess = await GeneratePMPTextFile(docPath, sFileName, dispatcher, uiSynchronizationContext, step);
                if (txtSuccess)
                {
                    await GeneratePMPExcelFromTemplate(docPath, sFileName, dispatcher, uiSynchronizationContext, step);
                }
            }
            catch (Exception ex)
            {
                AddLog(new LogEntry("ERROR", $"Erreur globale PMP : {ex.Message}"), dispatcher, uiSynchronizationContext);
                if (step != null) { step.Status = "Erreur"; step.ResultState = "Error"; }
            }
        }

        private string RemoveDiacritics(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;

            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder(capacity: normalizedString.Length);

            foreach (var c in normalizedString)
            {
                var unicodeCategory = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != System.Globalization.UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }

        private async Task<bool> GeneratePMPTextFile(string docPath, string sFileName, Dispatcher dispatcher, SynchronizationContext uiSynchronizationContext, WorkflowStep? step = null)
        {
            AddLog(new LogEntry("INFO", "Préparation de la génération PMP..."), dispatcher, uiSynchronizationContext);
            await Task.Delay(10); // Rendre la main à l'UI pour afficher le message initial

            try
            {
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
                            // 1. Remplacer les caractères accentués par des caractères non accentués
                            string noAccentLine = RemoveDiacritics(ligne);

                            // 2. Remplacer les caractères non alphanumériques sauf '/' par des espaces
                            string cleanedLine = regex.Replace(noAccentLine, " ");
                            
                            // Ajouter la ligne nettoyée à la ligne accumulée
                            accumulatedLine.Append(cleanedLine);
                        }
                    }

                    // Écrire des lignes de 293 caractères de longueur
                    // On soustrait 10 secondes pour forcer un premier affichage immédiat
                    DateTime lastLogTime = DateTime.Now.AddSeconds(-10); 
                    long linesWritten = 0;
                    StringBuilder finalOutput = new StringBuilder();

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

                            // Enregistrer la ligne dans le StringBuilder (en mémoire plutôt que sur le disque)
                            finalOutput.AppendLine(lineToWrite);
                            linesWritten++;

                            // Log "Traitement en cours ..." toutes les 10 secondes pour ne pas geler l'UI
                            if ((DateTime.Now - lastLogTime).TotalSeconds >= 10)
                            {
                                AddLog(new LogEntry("INFO", $"Formatage en cours ... {linesWritten} lignes préparées"), dispatcher, uiSynchronizationContext);
                                lastLogTime = DateTime.Now; // Réinitialiser le temps du dernier log
                                
                                // Rendre la main au thread UI pour que l'affichage puisse se mettre à jour
                                await Task.Delay(10);
                            }
                        }

                        AddLog(new LogEntry("INFO", "Sauvegarde du fichier texte sur le disque..."), dispatcher, uiSynchronizationContext);
                        
                        const int MaxRetries = 5; // Nombre maximal de tentatives
                        const int DelayMs = 1000; // Délai en cas d'occupation du fichier
                        for (int retry = 0; retry < MaxRetries; retry++)
                        {
                            try
                            {
                                await writer.WriteAsync(finalOutput.ToString());
                                break;
                            }
                            catch (IOException ex)
                            {
                                if (ex.HResult == -2147024864 && retry < MaxRetries - 1)
                                {
                                    AddLog(new LogEntry("WARNING", $"Le fichier est utilisé. Réessai {retry + 1}/{MaxRetries} dans {DelayMs}ms..."), dispatcher, uiSynchronizationContext);
                                    await Task.Delay(DelayMs);
                                }
                                else
                                {
                                    throw;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        AddLog(new LogEntry("ERROR", $"Erreur lors de la création du fichier PMP TXT : {ex.Message}"), dispatcher, uiSynchronizationContext);
                        if (step != null) { step.Status = "Erreur"; step.ResultState = "Error"; }
                        return false;
                    }
                }

                // Suppression de tous les fichiers sources traités
                foreach (string fichier in fichiersCSV)
                {
                    try
                    {
                        File.Delete(fichier);
                    }
                    catch (Exception exDelete)
                    {
                        AddLog(new LogEntry("WARNING", $"Impossible de supprimer le fichier {Path.GetFileName(fichier)} : {exDelete.Message}"), dispatcher, uiSynchronizationContext);
                    }
                }
                AddLog(new LogEntry("INFO", "Nettoyage terminé : " + fichiersCSV.Length + " fichier(s) source(s) supprimé(s)."), dispatcher, uiSynchronizationContext);

                AddLog(new LogEntry("SUCCESS", $"Fichiers PMP consolidés dans le dossier : " + docPath), dispatcher, uiSynchronizationContext);
                return true;

            }
            catch (Exception ex)
            {
                AddLog(new LogEntry("ERROR", $"Erreur globale lors de la génération TXT : {ex.Message}"), dispatcher, uiSynchronizationContext);
                if (step != null) { step.Status = "Erreur"; step.ResultState = "Error"; }
                return false;
            }
        }

        private async Task GeneratePMPExcelFromTemplate(string docPath, string sFileName, Dispatcher dispatcher, SynchronizationContext uiSynchronizationContext, WorkflowStep? step)
        {
            // --- INTEGRATION TEMPLATE EXCEL ---
            string sPMPExcelSaveAs = Path.Combine(docPath, $"PMPExcel_{DateTime.Now:yyMMddHHmmss}.xlsx");
            string pmpTxtFile = Path.Combine(docPath, sFileName);
            string templatePath = string.Empty;

            if (dispatcher != null)
            {
                dispatcher.Invoke(() =>
                {
                    Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog
                    {
                        Title = "Sélectionner le modèle Excel PMP",
                        Filter = "Fichiers Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Tous les fichiers (*.*)|*.*",
                        InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data")
                    };

                    if (openFileDialog.ShowDialog() == true)
                    {
                        templatePath = openFileDialog.FileName;
                    }
                });
            }

            if (string.IsNullOrEmpty(templatePath))
            {
                AddLog(new LogEntry("WARNING", "Génération annulée : aucun modèle PMP sélectionné."), dispatcher, uiSynchronizationContext);
                if (step != null) { step.Status = "Annulé"; step.ResultState = "Warning"; }
                return;
            }

            if (File.Exists(pmpTxtFile) && File.Exists(templatePath))
            {
                AddLog(new LogEntry("INFO", "Génération du modèle Excel à partir du fichier consolidé..."), dispatcher, uiSynchronizationContext);
                await Task.Run(() => 
                {
                    try 
                    {
                        using (var workbook = new XLWorkbook(templatePath))
                        {
                            var worksheet = workbook.Worksheet(1);
                            var pmpRange = worksheet.Range(8, 1, 10000, 22);
                            pmpRange.Clear();
                            pmpRange.Style.Fill.SetBackgroundColor(XLColor.White);
                            pmpRange.Style.Font.SetFontColor(XLColor.Black);
                            pmpRange.Style.Font.SetBold(false);
                            
                            string pattern = @"(?<!<[^<>]*)[^\w\s/](?![^<>]*>)";
                            Regex regexExcel = new Regex(pattern);
                            
                            string[] generatedLines = File.ReadAllLines(pmpTxtFile);
                            var dataList = new System.Collections.Generic.List<object[]>();
                            
                            foreach (string ligne in generatedLines)
                            {
                                if (ligne.Length >= 293)
                                {
                                    string cleanedLine = regexExcel.Replace(ligne, " ");
                                    string normalizedString = cleanedLine.Normalize(System.Text.NormalizationForm.FormD);
                                    StringBuilder sb = new StringBuilder();
                                    foreach (char c in normalizedString)
                                    {
                                        if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c) != System.Globalization.UnicodeCategory.NonSpacingMark)
                                            sb.Append(c);
                                    }
                                    string cleanedAccentLine = sb.ToString().ToUpper();
                                    
                                    string[] stringArray = new string[24];
                                    
                                    stringArray[0] = cleanedAccentLine.Substring(0, 39).Trim();
                                    stringArray[1] = cleanedAccentLine.Substring(40, 8).Trim();
                                    stringArray[2] = cleanedAccentLine.Substring(48, 4).Trim();
                                    stringArray[3] = cleanedAccentLine.Substring(52, 8).Trim();
                                    stringArray[4] = cleanedAccentLine.Substring(60, 4).Trim();
                                    stringArray[5] = cleanedAccentLine.Substring(64, 20).Trim();
                                    stringArray[6] = cleanedAccentLine.Substring(84, 20).Trim();
                                    
                                    string sExtract = cleanedAccentLine.Substring(104, 1).Trim();
                                    switch (sExtract) {
                                        case "1": stringArray[7] = "AHT"; break;
                                        case "2": stringArray[7] = "AST"; break;
                                        case "3": stringArray[7] = "MHP"; break;
                                        case "4": stringArray[7] = "MEP"; break;
                                        default: stringArray[7] = ""; break;
                                    }
                                    
                                    stringArray[8] = cleanedAccentLine.Substring(105, 60).Trim();
                                    stringArray[9] = cleanedAccentLine.Substring(165, 3).Trim();
                                    
                                    try {
                                        string t = cleanedAccentLine.Substring(168, 5).Trim();
                                        if (!string.IsNullOrEmpty(t)) stringArray[10] = TimeSpan.FromMinutes(int.Parse(t)).ToString(@"hh\:mm\:ss");
                                        else stringArray[10] = "";
                                    } catch { stringArray[10] = ""; }
                                    
                                    stringArray[11] = cleanedAccentLine.Substring(173, 3).Trim();
                                    
                                    try {
                                        string sExt = cleanedAccentLine.Substring(176, 2).Trim();
                                        switch(sExt) {
                                            case "1": stringArray[12] = "7"; break;
                                            case "2": stringArray[12] = "14"; break;
                                            case "4": stringArray[12] = "30"; break;
                                            case "8": stringArray[12] = "60"; break;
                                            case "12": case "3M": stringArray[12] = "90"; break;
                                            case "16": case "4M": stringArray[12] = "120"; break;
                                            case "24": case "6M": stringArray[12] = "180"; break;
                                            case "48": stringArray[12] = "360"; break;
                                            case "A1": stringArray[12] = "365"; break;
                                            case "A2": stringArray[12] = "730"; break;
                                            case "A3": stringArray[12] = "1095"; break;
                                            case "A4": stringArray[12] = "1460"; break;
                                            case "A5": stringArray[12] = "1825"; break;
                                            case "A6": stringArray[12] = "2190"; break;
                                            case "A9": stringArray[12] = "3285"; break;
                                            case "20J1": case "MJ": case "1E": case "3E": stringArray[12] = "1"; break;
                                            default: stringArray[12] = ""; break;
                                        }
                                    } catch { stringArray[12] = ""; }
                                    
                                    stringArray[13] = cleanedAccentLine.Substring(178, 6).Trim();
                                    stringArray[14] = cleanedAccentLine.Substring(184, 18).Trim();
                                    stringArray[15] = cleanedAccentLine.Substring(202, 13).Trim().Split(' ')[0];
                                    stringArray[16] = cleanedAccentLine.Substring(215, 40).Trim();
                                    stringArray[17] = cleanedAccentLine.Substring(255, 10).Trim();
                                    stringArray[18] = cleanedAccentLine.Substring(265, 1).Trim();
                                    
                                    if (stringArray[18] == "X") stringArray[19] = "";
                                    else { stringArray[18] = ""; stringArray[19] = "X"; }
                                    
                                    stringArray[20] = cleanedAccentLine.Substring(266, 25).Trim();
                                    stringArray[21] = string.IsNullOrEmpty(stringArray[20]) ? "N" : "O";
                                    stringArray[22] = cleanedAccentLine.Substring(291, 2).Trim().PadLeft(2, '0');
                                    stringArray[23] = "S";
                                    
                                    object[] rowData = new object[17];
                                    rowData[0] = stringArray[5];
                                    rowData[1] = stringArray[6];
                                    rowData[2] = stringArray[8];
                                    rowData[3] = stringArray[10];
                                    rowData[4] = stringArray[12];
                                    rowData[5] = stringArray[7];
                                    rowData[6] = stringArray[17];
                                    rowData[7] = stringArray[16];
                                    rowData[8] = stringArray[21];
                                    rowData[9] = stringArray[23];
                                    rowData[10] = stringArray[15];
                                    rowData[11] = stringArray[14];
                                    rowData[12] = stringArray[20];
                                    rowData[13] = $"{stringArray[1]}.{stringArray[22]}";
                                    rowData[14] = stringArray[18];
                                    rowData[15] = stringArray[19];
                                    rowData[16] = stringArray[3];
                                    
                                    dataList.Add(rowData);
                                }
                            }
                            
                            if (dataList.Count > 0)
                            {
                                worksheet.Cell(8, 1).InsertData(dataList);

                                // Mise en forme des cellules documentées (Colonnes A (1) à Q (17))
                                var dataRange = worksheet.Range(8, 1, 8 + dataList.Count - 1, 17);
                                dataRange.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                                dataRange.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
                                dataRange.Style.Font.SetFontName("Calibri");
                                dataRange.Style.Font.SetFontSize(9);
                            }
                            workbook.SaveAs(sPMPExcelSaveAs);
                        }
                        
                        AddLog(new LogEntry("SUCCESS", "Fichier Excel PMP généré : ", sPMPExcelSaveAs), dispatcher, uiSynchronizationContext);
                        try {
                            Process.Start(new ProcessStartInfo(sPMPExcelSaveAs) { UseShellExecute = true });
                        } catch { }

                        // Nettoyage: suppression du fichier PMP texte consolidé une fois l'Excel généré
                        try
                        {
                            if (File.Exists(pmpTxtFile))
                            {
                                File.Delete(pmpTxtFile);
                            }
                        }
                        catch (Exception exDelete)
                        {
                            AddLog(new LogEntry("WARNING", $"Impossible de supprimer le fichier texte consolidé : {exDelete.Message}"), dispatcher, uiSynchronizationContext);
                        }

                    }
                    catch (Exception innerEx)
                    {
                        AddLog(new LogEntry("ERROR", $"Erreur lors de la génération Excel PMP: {innerEx.Message}"), dispatcher, uiSynchronizationContext);
                        if (step != null) { step.Status = "Erreur"; step.ResultState = "Error"; }
                    }
                });
            }
            else if (!File.Exists(templatePath))
            {
                AddLog(new LogEntry("WARNING", $"Template Excel non trouvé dans : {templatePath}"), dispatcher, uiSynchronizationContext);
                if (step != null) { step.Status = "Erreur Modèle"; step.ResultState = "Error"; }
            }
            
            if (step != null && step.ResultState != "Error")
            {
                step.Status = "Terminé";
                step.ResultState = "Success";
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
