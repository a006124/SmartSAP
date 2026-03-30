using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows.Input;

namespace SmartSAP.ViewModels.Modules
{
    // Gammes : Extraction
    public class Module07ViewModel : ModuleDetailViewModelBase
    {
        public Module07ViewModel(MainViewModel mainViewModel, string title)
            : base(mainViewModel, title)
        {
        }

        public record ExcelColumnModel(
            string entete,
            string commentaires,
            string exemple,
            int longueurMaxi,
            IEnumerable<string>? valeursAutorisées,
            bool forcerMajuscule,
            bool forcerVide,
            bool forcerDocumentation,
            string règleDeGestion
        );

        protected override void InitializeSteps()
        {
            Steps = new ObservableCollection<WorkflowStep>
            {
                new WorkflowStep {
                    Title = "1. Saisie de la liste des gammes de maintenance à extraire de SAP",
                    Description = "Crée un nouveau fichier Excel modèle.",
                    Icon = "\xE70F",
                    ModuleStep = "M07-E1",
                    OpenFile = true,
                    ActionCommand = GenerateTemplateCommand
                },
                new WorkflowStep {
                    Title = "2. Extraction de SAP",
                    Description = "Exécute la transaction SAP 'ZP13'.",
                    Icon = "\xE768",
                    ModuleStep = "M07-E2",
                    ActionCommand = ExecuteSAPTransactionCommand
                },
                new WorkflowStep {
                    Title = "3. Reconstruction du fichier Excel PMP",
                    Description = "Reconstruit le fichier PMP à partir des fichiers TXT extraits de SAP.",
                    Icon = "\xE768",
                    ModuleStep = "M07-E3",
                    Parameters = {
                        new StepParameter("Générer un PMP Excel global", ParameterType.Boolean, false)
                        //new StepParameter("Commentaire d'extraction", ParameterType.Text, "Extraction automatique"),
                        //new StepParameter("Mode", ParameterType.Choice, "Normal", new string[] { "Normal", "Rapide", "Détaillé" })
                    },
                    ActionCommand = GeneratePMPExcelCommand
                }
            };
        }

        // GÉNÉRATION DU PMP EXCEL
        protected override async Task GeneratePMPExcel(WorkflowStep? step = null)
        {
            var uiSynchronizationContext = SynchronizationContext.Current;
            System.Windows.Threading.Dispatcher dispatcher = null;
            if (System.Windows.Application.Current != null)
                dispatcher = System.Windows.Application.Current.Dispatcher;

            if (step == null)
                step = Steps.FirstOrDefault(s => s.ActionCommand == GeneratePMPExcelCommand);

            if (step != null) step.ResultState = "Processing";

            // Lecture du paramètre 1 : "Générer un PMP Excel global"
            var step3 = Steps.FirstOrDefault(s => s.ModuleStep == "M07-E3");
            bool genererExcelGlobal = step3?.Parameters.Count > 0 && step3.Parameters[0].Value is true;

            string docPath = Path.GetDirectoryName(LastGeneratedExcelPath) ?? AppDomain.CurrentDomain.BaseDirectory;

            try
            {
                // ── ÉTAPE 1 : Récupérer tous les fichiers TXT sources (hors PMP_*) ──
                string[] sourceFiles = Directory.GetFiles(docPath, "*.txt")
                    .Where(f => !Path.GetFileName(f).StartsWith("PMP_"))
                    .ToArray();

                if (sourceFiles.Length == 0)
                {
                    AddLog(new LogEntry("WARNING", "Aucun fichier TXT source trouvé dans le dossier."), dispatcher, uiSynchronizationContext);
                    if (step != null) { step.Status = "Absent"; step.ResultState = "Error"; }
                    return;
                }

                int success = 0;
                int errors = 0;
                var tempPmpFiles = new System.Collections.Generic.List<string>(); // TXT intermédiaires conservés pour le global

                // ── ÉTAPE 2 : Un PMP Excel par fichier TXT ──────────────────────────
                foreach (string sourceFile in sourceFiles)
                {
                    string gamme = Path.GetFileNameWithoutExtension(sourceFile);
                    string tempName = $"PMP_{gamme}_{DateTime.Now:yyMMddHHmmss}.txt";

                    AddLog(new LogEntry("INFO", $"Traitement de la gamme : {gamme}..."), dispatcher, uiSynchronizationContext);

                    bool txtOk = await GeneratePMPTextFile(
                        docPath, tempName, dispatcher, uiSynchronizationContext, step,
                        overrideFiles: new[] { sourceFile });

                    if (txtOk)
                    {
                        string tempTxtPath = Path.Combine(docPath, tempName);
                        // Si global activé, on NE supprime PAS le TXT intermédiaire
                        await GeneratePMPExcelFromTemplate(docPath, tempName, dispatcher, uiSynchronizationContext, step,
                            deleteTxtAfter: !genererExcelGlobal);
                        if (genererExcelGlobal) tempPmpFiles.Add(tempTxtPath);
                        success++;
                    }
                    else
                    {
                        errors++;
                    }
                }

                // ── ÉTAPE 3 (optionnelle) : Excel global consolidé si paramètre activé ──
                if (genererExcelGlobal && tempPmpFiles.Count > 0)
                {
                    AddLog(new LogEntry("INFO", "Génération du PMP Excel global consolidé..."), dispatcher, uiSynchronizationContext);
                    string sFileNameGlobal = "PMP_GLOBAL_" + DateTime.Now.ToString("yyMMddHHmmss") + ".txt";

                    // Consolider les PMP TXT intermédiaires (encore présents car deleteTxtAfter=false)
                    bool globalOk = await GeneratePMPTextFile(
                        docPath, sFileNameGlobal, dispatcher, uiSynchronizationContext, step,
                        overrideFiles: tempPmpFiles); // GeneratePMPTextFile supprimera ces TXT intermédiaires

                    if (globalOk)
                        await GeneratePMPExcelFromTemplate(docPath, sFileNameGlobal, dispatcher, uiSynchronizationContext, step);
                }

                // ── Bilan final ─────────────────────────────────────────────────────
                if (errors == 0 && success > 0)
                {
                    AddLog(new LogEntry("SUCCESS", $"✓ {success} PMP Excel générés avec succès."), dispatcher, uiSynchronizationContext);
                    if (step != null) { step.Status = "Terminé"; step.ResultState = "Success"; }
                }
                else if (success > 0)
                {
                    AddLog(new LogEntry("WARNING", $"⚠ {success} succès / {errors} erreur(s)."), dispatcher, uiSynchronizationContext);
                    if (step != null) { step.Status = "Partiel"; step.ResultState = "Error"; }
                }
                else
                {
                    AddLog(new LogEntry("ERROR", $"✗ Aucun PMP Excel généré. {errors} erreur(s)."), dispatcher, uiSynchronizationContext);
                    if (step != null) { step.Status = "Erreur"; step.ResultState = "Error"; }
                }
            }
            catch (Exception ex)
            {
                AddLog(new LogEntry("ERROR", $"Erreur globale PMP : {ex.Message}"), dispatcher, uiSynchronizationContext);
                if (step != null) { step.Status = "Erreur"; step.ResultState = "Error"; }
            }
        }


        protected override async Task ExecuteSAPTransactionAsync(WorkflowStep? step = null)
        {
            if (step == null)
            {
                step = Steps.FirstOrDefault(s => s.ActionCommand == ExecuteSAPTransactionCommand);
            }

            //if (step != null && step.ResultState == "Error") return;

            try
            {
                // 1. Contrôle de la connexion SAP (Fusionné ici)
                Logs.Add(new LogEntry("INFO", "Contrôle de la connexion SAP..."));
                var connResult = await Task.Run(() => SAPManager.IsConnectedToSAP());

                // Mise à jour de la barre d'état globale
                MainViewModel.IsSAPConnected = connResult.IsSuccess;
                MainViewModel.SAPInstanceInfo = connResult.IsSuccess ? $"Instance : {connResult.InstanceInfo}" : "Non connecté";

                if (!connResult.IsSuccess)
                {
                    Logs.Add(new LogEntry("ERROR", connResult.ErrorMessage));
                    if (step != null) { step.Status = "Erreur Connexion"; step.ResultState = "Error"; }
                    return;
                }

                Logs.Add(new LogEntry("SUCCESS", $"✓ Connexion SAP OK : {connResult.InstanceInfo}"));

                // 2. Récupération de la session
                dynamic session = SAPManager.GetActiveSession();
                if (session == null)
                {
                    Logs.Add(new LogEntry("ERROR", "Impossible de récupérer une session SAP active."));
                    if (step != null) { step.Status = "Erreur session"; step.ResultState = "Error"; }
                    return;
                }



                Logs.Add(new LogEntry("INFO", "Lancement de la transaction ZP13..."));

                if (string.IsNullOrEmpty(LastGeneratedExcelPath) || !File.Exists(LastGeneratedExcelPath))
                {
                    Logs.Add(new LogEntry("ERROR", "Le fichier de données Excel est introuvable."));
                    if (step != null) { step.Status = "Erreur Fichier"; step.ResultState = "Error"; }
                    return;
                }

                int succesCount = 0;
                int errorCount = 0;
                string docPath = Path.GetDirectoryName(LastGeneratedExcelPath) ?? AppDomain.CurrentDomain.BaseDirectory;
                string LinesInError = string.Empty;

                try
                {
                    using (var workbook = new XLWorkbook(LastGeneratedExcelPath))
                    {
                        var worksheet = workbook.Worksheets.FirstOrDefault();
                        if (worksheet == null)
                        {
                            Logs.Add(new LogEntry("ERROR", "Le fichier Excel ne contient aucune feuille."));
                            if (step != null) { step.Status = "Erreur Fichier"; step.ResultState = "Error"; }
                            return;
                        }

                        // On commence à la ligne 2 pour ignorer l'en-tête
                        int rowCount = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                        for (int row = 2; row <= rowCount; row++)
                        {
                            // Récupérer la valeur de la colonne 1 (A) et colonne 2 (B)
                            string division = worksheet.Cell(row, 1).GetString().Trim();
                            string gamme = worksheet.Cell(row, 2).GetString().Trim();

                            // Si les deux colonnes sont vides, on ignore la ligne
                            if (string.IsNullOrWhiteSpace(division) && string.IsNullOrWhiteSpace(gamme)) continue;

                            string resultFile = string.Empty;
                            string result = await Task.Run(() => SAPManager.ExecuteZP13(session, division, gamme, docPath, out resultFile)); // Transaction SAP

                            var parts = result.Split('|');
                            if (parts.Length >= 2 && parts[1] == "OK")
                            {
                                succesCount++;
                                AddLog(new LogEntry("INFO", $"Ligne {row - 1}/{rowCount - 1} - Extraction réussie pour {division}/{gamme}."), System.Windows.Application.Current?.Dispatcher, SynchronizationContext.Current);
                            }
                            else if (parts.Length >= 2 && parts[1] == "NOK")
                            {
                                errorCount++;
                                string errMsg = parts.Length > 4 ? parts[4] : "Non précisée";
                                LinesInError+= $"{Environment.NewLine}'{division} {gamme}' : {errMsg}";
                                AddLog(new LogEntry("WARNING", $"Ligne {row - 1}/{rowCount - 1} - Extraction NOK pour {division}/{gamme}: {errMsg}"), System.Windows.Application.Current?.Dispatcher, SynchronizationContext.Current);
                            }
                            else
                            {
                                errorCount++;
                                string errMsg = parts.Length > 4 ? parts[4] : result;
                                LinesInError+= $"{Environment.NewLine}'{division} {gamme}' : {errMsg}";
                                AddLog(new LogEntry("ERROR", $"Ligne {row - 1}/{rowCount - 1} - Erreur pour {division}/{gamme}: {errMsg}"), System.Windows.Application.Current?.Dispatcher, SynchronizationContext.Current);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logs.Add(new LogEntry("ERROR", $"Erreur lors de la lecture du fichier Excel : {ex.Message}"));
                    if (step != null) { step.Status = "Erreur Lecture"; step.ResultState = "Error"; }
                    return;
                }

                if (errorCount == 0 && succesCount > 0)
                {
                    Logs.Add(new LogEntry("SUCCESS", $"✓ Terminé avec succès. {succesCount} ligne(s) traitée(s)."));
                    if (step != null) { step.Status = "Terminé"; step.ResultState = "Success"; }
                }
                else if (succesCount > 0 && errorCount > 0)
                {
                    Logs.Add(new LogEntry("WARNING", $"⚠ Terminé avec {errorCount} erreur(s) et {succesCount} succès.{Environment.NewLine}{LinesInError}"));
                    if (step != null) { step.Status = "Succès partiel"; step.ResultState = "Error"; }
                }
                else
                {
                    Logs.Add(new LogEntry("ERROR", $"✗ Aucune ligne traitée avec succès. {errorCount} erreur(s)."));
                    if (step != null) { step.Status = "Erreur SAP"; step.ResultState = "Error"; }
                }
            }
            catch (System.Exception ex)
            {
                Logs.Add(new LogEntry("ERROR", $"Erreur fatale lors de l'intégration SAP : {ex.Message}"));
                if (step != null) { step.Status = "Crash"; step.ResultState = "Error"; }
            }
        }

        // DÉFINITION DES COLONNES DE L'EXCEL MODELE
        protected override void InitializeExcelColumns(WorkflowStep? step = null)
        {
            ExcelColumns.Clear();

            // Chargement des données depuis JSON
            string dataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            if (!Directory.Exists(dataPath))
                dataPath = Path.Combine(Directory.GetCurrentDirectory(), "Data");

            var divisions = LoadJsonValues(Path.Combine(dataPath, "division.json"), "01-Division Localisation");

            var ExcelModel = new List<ExcelColumnModel>
            {
                // Entete - Commentaires - Données d'exemple - Longueur maxi - Valeurs autorisées - Majuscules forcées - Vide forcé - Documentation forcée - Règle de gestion
                new ("Division - 4 car (*)", "Documenter le code suivant les divisions gérées dans SAP", "MC02", 4, divisions, true, false, true, ""),
                new ("Gamme - 8 car (*)", "Documenter le code Gamme", "SMCP0001", 8, null, true, false, true, ""),
            };

            var columnsToAdd = ExcelModel.Select(d =>
                new Models.ExcelColumnDefinition(
                    entete: d.entete,
                    commentaires: d.commentaires,
                    exemple: d.exemple,
                    longueurMaxi: d.longueurMaxi,
                    valeursAutorisées: d.valeursAutorisées?.ToArray(),
                    forcerMajuscule: d.forcerMajuscule,
                    forcerVide: d.forcerVide,
                    forcerDocumentation: d.forcerDocumentation,
                    règleDeGestion: d.règleDeGestion
                )
            );

            foreach (var col in columnsToAdd)
            {
                ExcelColumns.Add(col);
            }


        }

        private string[] LoadJsonValues(string filePath, string propertyName)
        {
            try
            {
                if (!File.Exists(filePath)) return Array.Empty<string>();

                string jsonContent = File.ReadAllText(filePath);
                using var doc = JsonDocument.Parse(jsonContent);
                return doc.RootElement.EnumerateArray()
                    .Select(e => e.GetProperty(propertyName).GetString())
                    .Where(s => s != null)
                    .ToArray();
            }
            catch (Exception ex)
            {
                Logs.Add(new LogEntry("ERROR", $"Erreur lors du chargement de {filePath} : {ex.Message}"));
                return Array.Empty<string>();
            }
        }
    }
}
