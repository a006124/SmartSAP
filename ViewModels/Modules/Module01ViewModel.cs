using System.Collections.ObjectModel;

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
                new WorkflowStep { Title = "2. Contrôle et export des données", Description = "Contrôle et exporte les données (Format SAP).", Icon = "\xE762", ActionCommand = ExportFixedWidthCommand },
                new WorkflowStep { Title = "3. Intégration SAP", Description = "Exécute la transaction SAP pour créer les Postes Techniques.", Icon = "\xE8A5" },
                new WorkflowStep { Title = "4. Audit & Validation", Description = "Vérification la création des Postes Techniques et logs.", Icon = "\xE9A1" }
            };
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
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Code ABC - 1 car", "Indicateur de criticité ABC", "1", 1));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Code projet - 30 car", "Référence projet", "", 30));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Code produit - 30 car", "Référence produit", "", 30));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("A maintenir - 1 car", "Indicateur de maintenance (1=Oui)", "1", 1));
        }
    }
}
