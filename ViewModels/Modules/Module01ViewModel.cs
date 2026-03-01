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
                new WorkflowStep { Title = "1. Saisie des données de base", Description = "Crée un nouveau fichier Excel à renseigner à partir d'un modèle.", Icon = "\xE70F" },
                new WorkflowStep { Title = "2. Sélection du fichier de données", Description = "Sélectionne le fichier Excel contenant les données à charger.", Icon = "\xE762" },
                new WorkflowStep { Title = "3. Intégration SAP", Description = "Exécute la transaction SAP pour créer les Postes Techniques.", Icon = "\xE8A5" },
                new WorkflowStep { Title = "4. Audit & Validation", Description = "Vérification la création des Postes Techniques et logs.", Icon = "\xE9A1" }
            };
        }
        protected override void InitializeExcelColumns()
        {
            ExcelColumns.Add(new ExcelColumnDefinition("ID Equipement", "Numéro unique de l'équipement dans SAP", "EQUI-001"));
            ExcelColumns.Add(new ExcelColumnDefinition("Catégorie", "Code de catégorie d'équipement (ex: M pour Machine)", "M"));
            ExcelColumns.Add(new ExcelColumnDefinition("Fabricant", "Nom du constructeur de l'équipement", "SIEMENS"));
            ExcelColumns.Add(new ExcelColumnDefinition("Modèle", "Désignation précise du modèle", "SIMATIC S7-1200"));
        }
    }
}
