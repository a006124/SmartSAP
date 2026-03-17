using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;

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
                    Title = "2. Contrôle et export des données",
                    Description = "Contrôle et exporte les données (Format SAP). ",
                    Icon = "\xE762",
                    ModuleStep = "M01-E2",
                    NombreMini = 1,
                    OpenFile = false,
                    ActionCommand = ExportFixedWidthCommand
                },
                new WorkflowStep {
                    Title = "3. Intégration SAP",
                    Description = "Exécute la transaction SAP 'ZSMNBAO15'.",
                    Icon = "\xE768",
                    ModuleStep = "M01-E3",
                    ActionCommand = ExecuteSAPTransactionCommand
                }
            };
        }

        // EXÉCUTION DE LA TRANSACTION SAP
        protected override async Task ExecuteSAPTransactionAsync(WorkflowStep? step = null)
        {
            await base.ExecuteSAPTransactionAsync(step); // Vérifie la présence du fichier exporté

            if (step == null)
            {
                step = Steps.FirstOrDefault(s => s.ActionCommand == ExecuteSAPTransactionCommand);
            }

            if (step != null && step.ResultState == "Error") return;

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

                string resultFile = string.Empty;
                string result = await Task.Run(() => SAPManager.ExecuteZP13(session, LastExportedTextPath, out resultFile)); // Transaction SAP

                // Affichage du résultat brut dans les logs
                Logs.Add(new LogEntry("DEBUG", $"Réponse brute SAP : {result}"));

                var parts = result.Split('|');
                if (parts.Length >= 2 && parts[1] == "OK")
                {
                    Logs.Add(new LogEntry("SUCCESS", $"✓ Transaction terminée avec succès. Lignes lues: {parts[2]}."));
                    if (step != null) { step.Status = "Terminé"; step.ResultState = "Success"; }
                }
                else if (parts.Length >= 2 && parts[1] == "NOK")
                {
                    Logs.Add(new LogEntry("WARNING", $"⚠ Transaction terminée avec {parts[3]} erreur(s)."));
                    if (step != null) { step.Status = "Succès partiel"; step.ResultState = "Error"; }
                }
                else
                {
                    Logs.Add(new LogEntry("ERROR", $"✗ Erreur lors de l'exécution : {result}"));
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
