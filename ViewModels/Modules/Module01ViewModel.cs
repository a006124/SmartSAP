using System.Linq;
using System.IO;
using System.Text.Json;
using System;

namespace SmartSAP.ViewModels.Modules
{
    public class Module01ViewModel : ModuleDetailViewModelBase
    {
        public Module01ViewModel(MainViewModel mainViewModel, string title) 
            : base(mainViewModel, title)
        {
        }

        protected override void InitializeSteps()
        {
            Steps = new ObservableCollection<WorkflowStep>
            {
                new WorkflowStep { 
                    Title = "1. Saisie des données nécessaires à la création des Postes Techniques dans SAP", 
                    Description = "Crée un nouveau fichier Excel modèle.", 
                    Icon = "\xE70F", 
                    ModuleStep = "E1",
                    ActionCommand = GenerateTemplateCommand 
                },
                new WorkflowStep { 
                    Title = "2. Contrôle et export des données", 
                    Description = "Contrôle et exporte les données (Format SAP). ", 
                    Icon = "\xE762",
                    ModuleStep = "E2",
                    ActionCommand = ExportFixedWidthCommand 
                },
                new WorkflowStep { 
                    Title = "3. Intégration SAP", 
                    Description = "Exécute la transaction SAP 'ZSMNBAO15'.", 
                    Icon = "\xE768", 
                    ModuleStep = "E3",
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

                Logs.Add(new LogEntry("INFO", "Lancement de la transaction ZSMNBAO15..."));
                
                string resultFile = string.Empty;
                string result = await Task.Run(() => SAPManager.ExecuteZSMNBAO15(session, LastExportedTextPath, out resultFile));

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
        // Header - Commentaire - Données d'exemple - Largeur fixe - Majuscules forcées - Valeurs autorisées
        protected override void InitializeExcelColumns(WorkflowStep? step = null)
        {
            ExcelColumns.Clear();
            
            // Chargement des données depuis JSON
            string dataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            // Note: En mode Debug/Développement, le chemin peut varier, on essaie aussi le chemin relatif au projet
            if (!Directory.Exists(dataPath)) 
                dataPath = Path.Combine(Directory.GetCurrentDirectory(), "Data");

            var divisions = LoadJsonValues(Path.Combine(dataPath, "division.json"), "01-Division Localisation");
            var langues = LoadJsonValues(Path.Combine(dataPath, "langue.json"), "Langue préférée (division)");
            var abc = LoadJsonValues(Path.Combine(dataPath, "abc.json"), "abc");
            var a_maintenir = LoadJsonValues(Path.Combine(dataPath, "a_maintenir.json"), "a_maintenir");

            ExcelColumns.Add(new Models.ExcelColumnDefinition("Division - 4 car (*)", "Code Division", divisions.FirstOrDefault() ?? "MC02", 4, true, divisions));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Langue - 2 car (*)", "Code de langue (ex: FR)", langues.FirstOrDefault() ?? "FR", 2, true, langues));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Poste technique - 30 car (*)", "Nom du poste technique", "MC02_E_PT", 30, true));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Désignation - 40 car (*)", "Désignation de l'équipement", "PRESSE TRANSFERT", 40, true));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Localisation - 10 car", "Code de localisation", "150", 10, true));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Centre de coût - 10 car", "Code du centre de coût", "AC004510", 10, true));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Poste - 4 car", "Numéro de poste", "0010", 4, true));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Code ABC - 1 car", "Indicateur de criticité ABC", "1", 1, true, abc));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Code projet - 30 car", "Référence projet", "", 30, true));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Code produit - 30 car", "Référence produit", "", 30, true));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("A maintenir - 1 car", "Indicateur de maintenance (1=Oui)", "1", 1, true, a_maintenir));
        }

        private string[] LoadJsonValues(string filePath, string propertyName)
        {
            try
            {
                if (!File.Exists(filePath)) return Array.Empty<string>();

                string jsonContent = File.ReadAllText(filePath);
                using var doc = JsonDocument.Parse(jsonContent);
                return doc.RootElement.EnumerateArray()
                    .Select(e => e.GetProperty(propertyName).GetString() ?? "")
                    .Where(s => !string.IsNullOrEmpty(s))
                    .ToArray();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Erreur lors du chargement de {filePath} : {ex.Message}");
                return Array.Empty<string>();
            }
        }
    }
}
