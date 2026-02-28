using System.Collections.ObjectModel;
using System.Windows.Input;

namespace SmartSAP.ViewModels
{
    public class LibraryViewModel : ViewModelBase
    {
        private readonly MainViewModel _mainViewModel;

        public ObservableCollection<ModuleInfo> Modules { get; set; }
        public ICommand NavigateToModuleCommand { get; set; }

        public LibraryViewModel(MainViewModel mainViewModel)
        {
            _mainViewModel = mainViewModel;

            Modules = new ObservableCollection<ModuleInfo>
            {
                new ModuleInfo { Title = "Création d'Equipement",   Description = "Créer de nouveaux équipements dans SAP via Excel.", IconKind = "Plus",   Color = "#3B82F6" },
                new ModuleInfo { Title = "Modification d'Equipement", Description = "Modifier les équipements existants en masse.",       IconKind = "Pencil", Color = "#10B981" },
                new ModuleInfo { Title = "Suppression d'Equipement",  Description = "Archiver ou supprimer des équipements obsolètes.",   IconKind = "Trash",  Color = "#EF4444" }
            };

            NavigateToModuleCommand = new RelayCommand(p => _mainViewModel.NavigateToModule((string)p!));
        }
    }

    public class ModuleInfo
    {
        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string IconKind { get; set; } = string.Empty;
        public string Color { get; set; } = string.Empty;
        public string Version { get; set; } = "v1.0.0";
        public string HealthStatus { get; set; } = "Optimal";
    }
}
