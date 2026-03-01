using System.Collections.ObjectModel;

namespace SmartSAP.ViewModels.Modules
{
    public class Module05ViewModel : ModuleDetailViewModelBase
    {
        public Module05ViewModel(MainViewModel mainViewModel, string title) 
            : base(mainViewModel, title)
        {
            InitializeSteps();
            CompleteInitialization();
        }

        protected override void InitializeSteps()
        {
            Steps = new ObservableCollection<WorkflowStep>
            {
                new WorkflowStep { Title = "1. Saisie des donnÃ©es de base", Description = "Renseigner l'identification, la catÃ©gorie et le constructeur dans le modÃ¨le Excel.", Icon = "ç0F" },
                new WorkflowStep { 
                    Title = "2. Données d'organisation", 
                    Description = "Affecter le centre de coûts, l'entreprise et les domaines d'activité. ", 
                    Icon = "\xE762",
                    LinkText = "modifier le fichier",
                    LinkCommand = PickExcelFileCommand
                },
                new WorkflowStep { Title = "3. IntÃ©gration SAP (BAPI)", Description = "Appel de la BAPI_EQUI_CREATE pour gÃ©nÃ©rer les Ã©quipements.", Icon = "èA5" },
                new WorkflowStep { Title = "4. Audit & Validation", Description = "VÃ©rification des numÃ©ros d'Ã©quipements gÃ©nÃ©rÃ©s et logs.", Icon = "éA1" }
            };
        }
    }
}
