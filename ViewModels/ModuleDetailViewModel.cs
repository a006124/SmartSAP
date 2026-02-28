using System.Collections.ObjectModel;
using System.Windows.Input;

namespace SmartSAP.ViewModels
{
    public class ModuleDetailViewModel : ViewModelBase
    {
        private readonly MainViewModel _mainViewModel;

        public string ModuleTitle { get; set; }
        public ObservableCollection<WorkflowStep> Steps { get; set; }
        public ObservableCollection<LogEntry> Logs { get; set; }
        public ICommand GoBackCommand { get; set; }
        public ICommand RunWorkflowCommand { get; set; }

        public ModuleDetailViewModel(MainViewModel mainViewModel, string title)
        {
            _mainViewModel = mainViewModel;
            ModuleTitle = title;

            Logs = new ObservableCollection<LogEntry>();
            Steps = new ObservableCollection<WorkflowStep>
            {
                new WorkflowStep { Title = "Création Modèle Excel",  Description = "Générer le template Excel standard.",          Icon = "", Status = "Ready" },
                new WorkflowStep { Title = "Sélection Fichier Excel", Description = "Choisir le fichier contenant les données.",     Icon = "", Status = "Ready" },
                new WorkflowStep { Title = "Exécution SAP",           Description = "Lancer l'intégration dans SAP.",               Icon = "", Status = "Ready" },
                new WorkflowStep { Title = "Visualisation Résultats", Description = "Vérifier le statut du traitement.",            Icon = "", Status = "Ready" }
            };

            GoBackCommand = new RelayCommand(_ => _mainViewModel.NavigateToLibrary());
            RunWorkflowCommand = new RelayCommand(async _ => await ExecuteWorkflowAsync());
        }

        private async Task ExecuteWorkflowAsync()
        {
            Logs.Add(new LogEntry("INFO", $"Démarrage du workflow : {ModuleTitle}"));

            foreach (var step in Steps)
            {
                step.Status = "Processing";
                Logs.Add(new LogEntry("INFO", $"Exécution de : {step.Title}"));

                await Task.Delay(1500);

                step.Status = "Completed";
                Logs.Add(new LogEntry("SUCCESS", $"{step.Title} terminé avec succès."));
            }

            Logs.Add(new LogEntry("INFO", "Workflow terminé."));
        }
    }

    public class WorkflowStep : ViewModelBase
    {
        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Icon { get; set; } = string.Empty;

        private string _status = string.Empty;
        public string Status
        {
            get => _status;
            set => SetProperty(ref _status, value);
        }
    }

    public class LogEntry
    {
        public string Timestamp { get; } = DateTime.Now.ToString("HH:mm:ss");
        public string Type { get; set; }
        public string Message { get; set; }

        public LogEntry(string type, string message)
        {
            Type = type;
            Message = message;
        }
    }
}
