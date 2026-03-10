using System.Collections.ObjectModel;

namespace SmartSAP.ViewModels.Modules
{
    public class Module02ViewModel : ModuleDetailViewModelBase
    {
        // Poste Technique : Modification en masse
        public Module02ViewModel(MainViewModel mainViewModel, string title) 
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
                    Description = "Affecter le centre de coûts, l'entreprise et les domaines d'activitÃ©. ", 
                    Icon = "\xE762",
                    LinkText = "modifier le fichier",
                    LinkCommand = PickExcelFileCommand
                },
                new WorkflowStep { Title = "3. Intégration SAP (BAPI)", Description = "Appel de la BAPI_EQUI_CREATE pour générer les équipements.", Icon = "\xE8A5", ActionCommand = CheckSAPConnectionCommand },
                new WorkflowStep { Title = "4. Audit & Validation", Description = "VÃ©rification des numÃ©ros d'Ã©quipements gÃ©nÃ©rÃ©s et logs.", Icon = "éA1" }
            };
        }
    }
}
