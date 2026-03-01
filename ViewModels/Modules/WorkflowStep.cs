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
    }
}
