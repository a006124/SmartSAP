using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows.Input;

namespace SmartSAP.ViewModels.Modules
{
    public abstract class ModuleDetailViewModelBase : ViewModelBase
    {
        protected readonly MainViewModel MainViewModel;

        public string ModuleTitle { get; protected set; }
        public ObservableCollection<WorkflowStep> Steps { get; protected set; }
        public ObservableCollection<LogEntry> Logs { get; protected set; }
        public ICommand GoBackCommand { get; protected set; }
        public ICommand RunWorkflowCommand { get; protected set; }

        protected ModuleDetailViewModelBase(MainViewModel mainViewModel, string title)
        {
            MainViewModel = mainViewModel;
            ModuleTitle = title;
            
            Logs = new ObservableCollection<LogEntry>();
            Steps = new ObservableCollection<WorkflowStep>();

            GoBackCommand = new RelayCommand(_ => MainViewModel.NavigateToLibrary());
            RunWorkflowCommand = new RelayCommand(async _ => await ExecuteWorkflowAsync());
        }

        protected virtual void InitializeSteps()
        {
            // A surcharger dans les classes enfants pour définir les étapes spécifiques
        }

        protected virtual async Task ExecuteWorkflowAsync()
        {
            Logs.Add(new LogEntry("INFO", $"Démarrage du workflow : {ModuleTitle}"));

            foreach (var step in Steps)
            {
                step.Status = "Processing";
                Logs.Add(new LogEntry("INFO", $"Exécution de : {step.Title}"));

                await Task.Delay(1500); // Simulation

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

        private string _status = "Ready";
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
