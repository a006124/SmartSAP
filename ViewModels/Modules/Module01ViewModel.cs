using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Linq;

namespace SmartSAP.ViewModels.Modules
{
    public class Module01ViewModel : ModuleDetailViewModelBase
    {
        public Module01ViewModel(MainViewModel mainViewModel, string title) 
            : base(mainViewModel, title)
        {
            InitializeSteps();
            InitializeExcelColumns();
            CompleteInitialization();
        }

        protected override void InitializeSteps()
        {
            Steps = new ObservableCollection<WorkflowStep>
            {
                new WorkflowStep { Title = "1. Saisie des données de base", Description = "Crée un nouveau fichier Excel à renseigner à partir d'un modèle.", Icon = "\xE70F", ActionCommand = GenerateTemplateCommand },
                new WorkflowStep { 
                    Title = "2. Contrôle et export des données", 
                    Description = "Contrôle et exporte les données (Format SAP). ", 
                    Icon = "\xE762", 
                    ActionCommand = ExportFixedWidthCommand
                },
                new WorkflowStep { Title = "3. Connexion SAP", Description = "Vérifie la connexion au serveur SAP.", Icon = "\xE8A5", ActionCommand = CheckSAPConnectionCommand },
                new WorkflowStep { Title = "4. Intégration SAP", Description = "Exécute la transaction ZSMNBAO15.", Icon = "\xE768", ActionCommand = ExecuteSAPTransactionCommand }
            };
        }

        protected override async Task ExecuteSAPTransactionAsync()
        {
            await base.ExecuteSAPTransactionAsync(); // Vérifie le fichier
            
            var step = Steps.FirstOrDefault(s => s.ActionCommand == ExecuteSAPTransactionCommand);
            if (step != null && step.ResultState == "Error") return; // Arrêt si fichier absent (déjà loggé par base)

            try
            {
                dynamic session = SAPManager.GetActiveSession();
                if (session == null)
                {
                    Logs.Add(new LogEntry("ERROR", "Impossible de récupérer une session SAP active. Veuillez vérifier l'étape 3."));
                    if (step != null) { step.Status = "Erreur session"; step.ResultState = "Error"; }
                    return;
                }

                Logs.Add(new LogEntry("INFO", "Lancement de la transaction ZSMNBAO15..."));
                
                string resultFile = string.Empty;
                string result = await Task.Run(() => SAPManager.ExecuteZSMNBAO15(session, LastExportedTextPath, out resultFile));

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

                if (!string.IsNullOrEmpty(resultFile) && System.IO.File.Exists(resultFile))
                {
                    Logs.Add(new LogEntry("SUCCESS", "Le fichier de log SAP a été généré : ", resultFile));
                    Process.Start(new ProcessStartInfo(resultFile) { UseShellExecute = true });
                }
            }
            catch (System.Exception ex)
            {
                Logs.Add(new LogEntry("ERROR", $"Erreur fatale lors de l'intégration SAP : {ex.Message}"));
                if (step != null) { step.Status = "Crash"; step.ResultState = "Error"; }
            }
        }

        protected override void InitializeExcelColumns()
        {
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Division - 4 car (*)", "Numéro unique de l'équipement dans SAP", "MC02", 4));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Langue - 2 car (*)", "Code de langue (ex: FR)", "FR", 2));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Poste technique - 30 car (*)", "Nom du poste technique", "MC02_E_PT", 30));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Désignation - 40 car (*)", "Désignation de l'équipement", "PRESSE TRANSFERT", 40));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Localisation - 10 car", "Code de localisation", "150", 10));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Centre de coût - 10 car", "Code du centre de coût", "AC004510", 10));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Poste - 4 car", "Numéro de poste", "0010", 4));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Code ABC - 1 car", "Indicateur de criticité ABC", "1", 1, true, new[] { "1", "2", "3" }));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Code projet - 30 car", "Référence projet", "", 30));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Code produit - 30 car", "Référence produit", "", 30));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("A maintenir - 1 car", "Indicateur de maintenance (1=Oui)", "1", 1, true, new[] { "0", "1" }));
        }
    }
}
