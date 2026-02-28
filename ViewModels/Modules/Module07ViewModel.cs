using System.Collections.ObjectModel;

namespace SmartSAP.ViewModels.Modules
{
    public class Module07ViewModel : ModuleDetailViewModelBase
    {
        public Module07ViewModel(MainViewModel mainViewModel, string title) 
            : base(mainViewModel, title)
        {
            InitializeSteps();
        }

        protected override void InitializeSteps()
        {
            Steps = new ObservableCollection<WorkflowStep>
            {
                new WorkflowStep { Title = "1. Saisie des données de base", Description = "Renseigner l'identification, la catégorie et le constructeur dans le modèle Excel.", Icon = "�0F" },
                new WorkflowStep { Title = "2. Données d'organisation", Description = "Affecter le centre de coûts, l'entreprise et les domaines d'activité.", Icon = "�62" },
                new WorkflowStep { Title = "3. Intégration SAP (BAPI)", Description = "Appel de la BAPI_EQUI_CREATE pour générer les équipements.", Icon = "�A5" },
                new WorkflowStep { Title = "4. Audit & Validation", Description = "Vérification des numéros d'équipements générés et logs.", Icon = "�A1" }
            };
        }
    }
}
