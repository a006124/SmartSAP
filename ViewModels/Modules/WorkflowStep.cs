using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using SmartSAP.ViewModels;

namespace SmartSAP.ViewModels.Modules
{
    public class WorkflowStep : ViewModelBase
    {
        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Icon { get; set; } = string.Empty;
        public string ModuleStep { get; set; } = string.Empty;
        public int NombreMini { get; set; } = 1; // Nombre minimum de lignes nécessaires
        public bool OpenFile { get; set; } = false;

        private string _status = "Ready";
        public string Status
        {
            get => _status;
            set => SetProperty(ref _status, value);
        }

        private string _resultState = "Normal";
        public string ResultState
        {
            get => _resultState;
            set => SetProperty(ref _resultState, value);
        }

        public string? LinkText { get; set; }
        public ICommand? LinkCommand { get; set; }

        private bool _isLast;
        public bool IsLast
        {
            get => _isLast;
            set => SetProperty(ref _isLast, value);
        }

        public ICommand? ActionCommand { get; set; }

        public ObservableCollection<StepParameter> Parameters { get; } = new ObservableCollection<StepParameter>();

        public bool HasSettings
        {
            get => Parameters.Any();
        }

        private bool _isSettingsOpen;
        public bool IsSettingsOpen
        {
            get => _isSettingsOpen;
            set => SetProperty(ref _isSettingsOpen, value);
        }

        public ICommand ToggleSettingsCommand { get; }
        public ICommand ConfirmSettingsCommand { get; }
        public ICommand CancelSettingsCommand { get; }

        // Snapshot des valeurs pour restauration en cas d'annulation
        private Dictionary<StepParameter, object?> _parameterSnapshot = new();

        public WorkflowStep()
        {
            ToggleSettingsCommand = new RelayCommand(o =>
            {
                if (!IsSettingsOpen)
                {
                    // Sauvegarde des valeurs courantes avant ouverture
                    _parameterSnapshot = Parameters.ToDictionary(p => p, p => p.Value);
                }
                IsSettingsOpen = !IsSettingsOpen;
            });

            ConfirmSettingsCommand = new RelayCommand(o =>
            {
                _parameterSnapshot.Clear();
                IsSettingsOpen = false;
            });

            CancelSettingsCommand = new RelayCommand(o =>
            {
                // Restauration des valeurs d'origine
                foreach (var (param, originalValue) in _parameterSnapshot)
                    param.Value = originalValue;
                _parameterSnapshot.Clear();
                IsSettingsOpen = false;
            });

            Parameters.CollectionChanged += (s, e) => OnPropertyChanged(nameof(HasSettings));
        }
    }
}
