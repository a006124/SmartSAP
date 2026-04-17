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
    // FID : Modification
    public class Module11ViewModel : ModuleDetailViewModelBase
    {
        public Module11ViewModel(MainViewModel mainViewModel, string title)
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
                    Title = "1. Saisie de la liste des FID à contrôler / modifier dans SAP",
                    Description = "Crée un nouveau fichier Excel modèle.",
                    Icon = "\xE70F",
                    ModuleStep = "M11-E1",
                    OpenFile = true,
                    ActionCommand = GenerateTemplateCommand
                },
                new WorkflowStep {
                    Title = "2. Exécution de la transaction SAP",
                    Description = "Exécute la transaction SAP 'CV02N'.",
                    Icon = "\xE768",
                    ModuleStep = "M11-E2",                  
                    ActionCommand = ExecuteSAPTransactionCommand
                }
            };
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



                Logs.Add(new LogEntry("INFO", "Lancement de la transaction CV03N..."));

                if (string.IsNullOrEmpty(LastGeneratedExcelPath) || !File.Exists(LastGeneratedExcelPath))
                {
                    Logs.Add(new LogEntry("ERROR", "Le fichier de données Excel est introuvable."));
                    if (step != null) { step.Status = "Erreur Fichier"; step.ResultState = "Error"; }
                    return;
                }

                int succesCount = 0;
                int errorCount = 0;
                // Le fichier est maintenant dans "Fichiers Temporaires", donc le dossier racine est le parent
                string currentParent = Path.GetDirectoryName(LastGeneratedExcelPath) ?? AppDomain.CurrentDomain.BaseDirectory;
                string baseDir = Path.GetFileName(currentParent) == "Fichiers Temporaires" 
                    ? Path.GetDirectoryName(currentParent) 
                    : currentParent;
                
                string docPath = Path.Combine(baseDir, "Fichiers Source");
                
                // S'assurer que le dossier existe au cas où
                if (!Directory.Exists(docPath)) Directory.CreateDirectory(docPath);
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
                        List<string> tabResult = new List<string>();
                        for (int row = 2; row <= rowCount; row++)
                        {
                            // Récupérer la valeur de la colonne 1 (A) et colonne 2 (B)
                            string fidCode = "S" + worksheet.Cell(row, 4).GetString().Trim();
                            string fidType = worksheet.Cell(row, 3).GetString().Trim();
                            string fidLibellé = worksheet.Cell(row, 7).GetString().Trim();
                            string fidStatut = worksheet.Cell(row, 17).GetString().Trim();

                            // Si les deux colonnes sont vides, on ignore la ligne
                            if (string.IsNullOrWhiteSpace(fidCode)) continue;

                            string resultFile = string.Empty;
                            string result = await Task.Run(() => SAPManager.ExecuteCV03NTDI(session, fidCode, fidType, fidLibellé, fidStatut, out resultFile)); // Transaction SAP

                            var parts = result.Split('|');
                            if (parts.Length >= 2 && parts[1] == "OK")
                            {
                                succesCount++;
                                AddLog(new LogEntry("INFO", $"Ligne {row - 1}/{rowCount - 1} - Contrôle OK de la FID n°{fidCode + "_" + fidType}."), System.Windows.Application.Current?.Dispatcher, SynchronizationContext.Current);
                                tabResult.Add(result);
                            }
                            else if (parts.Length >= 2 && parts[1] == "NOK")
                            {
                                errorCount++;
                                string errMsg = parts.Length > 4 ? parts[4] : "Non précisée";
                                LinesInError+= $"{Environment.NewLine}'{fidCode + "_" + fidType}' : {errMsg}";
                                AddLog(new LogEntry("WARNING", $"Ligne {row - 1}/{rowCount - 1} - Contrôle NOK de la FID n°{fidCode + "_" + fidType}: {errMsg}"), System.Windows.Application.Current?.Dispatcher, SynchronizationContext.Current);
                                tabResult.Add(result);
                            }
                            else
                            {
                                errorCount++;
                                string errMsg = parts.Length > 4 ? parts[4] : result;
                                LinesInError+= $"{Environment.NewLine}'{fidCode + "_" + fidType}' : {errMsg}";
                                AddLog(new LogEntry("ERROR", $"Ligne {row - 1}/{rowCount - 1} - Erreur sur la FID n°{fidCode + "_" + fidType}: {errMsg}"), System.Windows.Application.Current?.Dispatcher, SynchronizationContext.Current);
                                tabResult.Add(string.Empty);
                            }
                        }


                        // Traitement du fichier Excel 
//                        if (!string.IsNullOrEmpty(LastGeneratedSAPExcelPath) && System.IO.File.Exists(LastGeneratedSAPExcelPath))
//                        {
                            // Exécution de la fonction EnrichirFromSAPExcelWorkbookM11_E_2
                            try
                            {
                                var excelService = new SmartSAP.Services.Excel.ExcelManager();
                                string enrichResult = excelService.EnrichirFromSAPExcelWorkbookM11_E_2(LastGeneratedExcelPath, tabResult);
                                Logs.Add(new LogEntry("SUCCESS", $"Enrichissement terminé : {enrichResult}"));
                            }
                            catch (System.Exception ex)
                            {
                                Logs.Add(new LogEntry("ERROR", $"Erreur lors de l'enrichissement : {ex.Message}"));
                            }
//                        }
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

            var type_FID = LoadJsonValues(Path.Combine(dataPath, "type_FID.json"), "type_FID");
            var langue = LoadJsonValues(Path.Combine(dataPath, "langue.json"), "Langue préférée (division)");
            var statut_FID= LoadJsonValues(Path.Combine(dataPath, "statut_FID.json"), "statut_FID");
            var type_FID_R = LoadJsonValues(Path.Combine(dataPath, "type_FID_R.json"), "type_FID_R");
            var version_FID = LoadJsonValues(Path.Combine(dataPath, "version_FID.json"), "version_FID");
            var action_lien_FID = LoadJsonValues(Path.Combine(dataPath, "action_lien_FID.json"), "action_lien_FID");
            var type_objet_lien_FID = LoadJsonValues(Path.Combine(dataPath, "type_objet_lien_FID.json"), "type_objet_lien_FID");
            var action_original_FID = LoadJsonValues(Path.Combine(dataPath, "action_original_FID.json"), "action_original_FID");
            var application_original_FID = LoadJsonValues(Path.Combine(dataPath, "application_original_FID.json"), "application_original_FID");



            var ExcelModel = new List<ExcelColumnModel>
            {
                // Entete - Commentaires - Données d'exemple - Longueur maxi - Valeurs autorisées - Majuscules forcées - Vide forcé - Documentation forcée - Règle de gestion
                new ("Type", "", "", 100, null, true, false, false, ""),
                new ("Message", "", "", 100, null, true, false, false, ""),
                new ("Type *", "", "", 100, null, true, false, false, ""),
                new ("Nom * SANS LE S DEVANT", "", "", 100, null, true, false, false, ""),
                new ("Indice (- si vide)", "", "", 100, null, true, false, false, ""),
                new ("Code langue *", "", "", 100, null, true, false, false, ""),
                new ("Libellé *", "", "", 100, null, true, false, false, ""),
                new ("FID R père SANS LE S DEVANT", "", "", 100, null, true, false, false, ""),
                new ("Observations (libellé long dans SAP)", "", "", 100, null, true, false, false, ""),
                new ("Liens EQT à ajouter", "", "", 100, null, true, false, false, ""),
                new ("Liens PT à ajouter", "", "", 100, null, true, false, false, ""),
                new ("Liens Article à ajouter", "", "", 100, null, true, false, false, ""),
                new ("Plan référence SANS LE S DEVANT", "", "", 100, null, true, false, false, ""),
                new ("Chemin des fichiers *", "", "", 100, null, true, false, false, ""),
                new ("Fichiers *", "", "", 100, null, true, false, false, ""),
                new ("Indice", "", "", 100, null, true, false, false, ""),
                new ("Statut", "", "", 100, null, true, false, false, ""),
                new ("Nombre de fichiers", "", "", 100, null, true, false, false, ""),
                new ("", "", "", 100, null, true, false, false, ""),
                new ("FID existante ?", "", "", 100, null, true, false, false, ""),
                new ("Version", "", "", 100, null, true, false, false, ""),
                new ("Description", "", "", 100, null, true, false, false, ""),
                new ("Statut initial", "", "", 100, null, true, false, false, ""),
                new ("Liaison Article (Code)", "", "", 100, null, true, false, false, ""),
                new ("Liaison Article (Désignation)", "", "", 100, null, true, false, false, ""),
                new ("Liaison Equipement (Code)", "", "", 100, null, true, false, false, ""),
                new ("Liaison Equipement (Désignation)", "", "", 100, null, true, false, false, ""),
                new ("Liaison Poste Technique (Code)", "", "", 100, null, true, false, false, ""),
                new ("Liaison Poste Technique (Désignation)", "", "", 100, null, true, false, false, ""),
                new ("Originaux (Nombre)", "", "", 100, null, true, false, false, ""),
                new ("Originaux (Fichier)", "", "", 100, null, true, false, false, ""),
                new ("Originaux (Cohérent ?)", "", "", 100, null, true, false, false, ""),
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
