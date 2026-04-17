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
    // Gamme : Créer gamme
    public class Module12ViewModel : ModuleDetailViewModelBase
    {
        public Module12ViewModel(MainViewModel mainViewModel, string title)
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
                    Title = "1. Saisie de la liste des gammes de maintenance à créer",
                    Description = "Crée un nouveau fichier Excel modèle.",
                    Icon = "\xE70F",
                    ModuleStep = "M12-E1",
                    OpenFile = true,
                    ActionCommand = GenerateTemplateCommand
                },
                new WorkflowStep {
                    Title = "2. Exécution SAP",
                    Description = "Exécute la transaction SAP 'IW32'.",
                    Icon = "\xE768",
                    ModuleStep = "M12-E2",
                    ActionCommand = ExecuteSAPTransactionCommand
                }
            };
        }

        // EXÉCUTION DE LA TRANSACTION SAP
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



                Logs.Add(new LogEntry("INFO", "Lancement de la transaction IW32..."));

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
                            // Récupérer la valeur de la colonne 1 (A) 
                            string OT = worksheet.Cell(row, 1).GetString().Trim();

                            // Si la colonne est vide, on ignore la ligne
                            if (string.IsNullOrWhiteSpace(OT)) continue;

                            string resultFile = string.Empty;
                            string result = await Task.Run(() => SAPManager.ExecuteIW32(session, OT, out resultFile)); // Transaction SAP

                            var parts = result.Split('|');
                            if (parts.Length >= 2 && parts[1] == "OK")
                            {
                                succesCount++;
                            }
                            else if (parts.Length >= 2 && parts[1] == "NOK")
                            {
                                errorCount++;
                                LinesInError += $"{Environment.NewLine}'{OT}' : {parts[4]}";
                            }
                            else
                            {
                                errorCount++;
                                LinesInError += $"{Environment.NewLine}'{OT}' : {parts[4]}";
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
            var gamme_utilisation = LoadJsonValues(Path.Combine(dataPath, "gamme_utilisation.json"), "gamme_utilisation");
            var gamme_statut = LoadJsonValues(Path.Combine(dataPath, "gamme_statut.json"), "gamme_statut");
            var gamme_etat_immo = LoadJsonValues(Path.Combine(dataPath, "gamme_etat_immo.json"), "gamme_etat_immo");

            var ExcelModel = new List<ExcelColumnModel>
            {
                // Entete - Commentaires - Données d'exemple - Longueur maxi - Valeurs autorisées - Majuscules forcées - Vide forcé - Documentation forcée - Règle de gestion
                new ("Groupe de gammes (*) - 8 car", "Documenter le groupe de gammes", "SDOE0001", 8, null, true, false, true, ""), // règle de gestion : doit débuter par SM2E et être suivi de 4 chiffres
                new ("Création Groupe de gammes (*) - 1 car", "Créer le groupe de gammes (1 = Oui / 0 = Non)", "1", 1, null, true, false, true, ""), // règle de gestion : doit débuter par SM2E et être suivi de 4 chiffres
                new ("Cpteur grpes gammes (*) - 2 car", "Documenter le compteur de groupe de gammes", "1", 2, null, true, false, true, ""), // règle de gestion : doit être un nombre entre 1 et 99
                new ("Désignation de la gamme (*) - 40 car", "Documenter la désignation de la gamme", "", 40, null, true, false, true, ""),
                new ("Division de planif. (*) - 4 car", "Documenter le code de la division de planification (ex : DO01)", "DO01", 4, null, true, false, true, ""),
                new ("Poste de travail (*)", "Documenter le code du poste de travail (ex : ST)", "ST", 8, null, true, false, true, ""),
                new ("Division (*) - 4 car", "Documenter le code de la division (ex : DO01)", "DO01", 4, null, true, false, true, ""),
                new ("Utilisation (*) - 2 car", "Documenter l'utilisation de la gamme (ex : 4 = Maintenance)", "4", 2, gamme_utilisation, true, false, true, ""),
                new ("Groupe de gestionnaire - 3 car", "Documenter le groupe de gestionnaire", "", 3, null, true, false, false, ""),
                new ("Statut de la gamme (*) - 1 car", "Documenter le statut de la gamme (ex : 2 = OK pour ordre)", "2", 1, gamme_statut, true, false, true, ""),
                new ("Etat immo. (*) - 1 car", "Documenter l'état d'immobilisation (1 = AHT, 2 = AST, 3 = MEP, 4 = MHP)", "2", 1, gamme_etat_immo, true, false, true, ""),
                new ("Stratégie d'entret. (*) - 6 car", "Documenter la stratégie d'entretien (ex : DOE12N)", "DOE12N", 6, null, true, false, true, ""),
                // new ("Sous-ensemble - 20 car", "Documenter le sous-ensemble", "", 20, null, true, false, true, ""),
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
