using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;

namespace SmartSAP.ViewModels
{
    public class LibraryViewModel : ViewModelBase
    {
        private readonly MainViewModel _mainViewModel;

        public ObservableCollection<ModuleInfo> Modules { get; set; }
        public ICommand NavigateToModuleCommand { get; set; }
        
        private ObservableCollection<ModuleInfo> _filteredModules;
        public ObservableCollection<ModuleInfo> FilteredModules
        {
            get => _filteredModules;
            set
            {
                _filteredModules = value;
                OnPropertyChanged(nameof(FilteredModules));
            }
        }
        
        private string _searchQuery = "";
        public string SearchQuery
        {
            get => _searchQuery;
            set
            {
                if (_searchQuery != value)
                {
                    _searchQuery = value;
                    OnPropertyChanged(nameof(SearchQuery));
                    FilterModules();
                }
            }
        }

        public LibraryViewModel(MainViewModel mainViewModel)
        {
            _mainViewModel = mainViewModel;

            Modules = new ObservableCollection<ModuleInfo>
            {
                new ModuleInfo { Title = "Création d'Equipement",   Description = "Créer de nouveaux équipements dans SAP via Excel.", IconKind = "Plus",   Color = "#3B82F6" },
                new ModuleInfo { Title = "Modification d'Equipement", Description = "Modifier les équipements existants en masse.",       IconKind = "Pencil", Color = "#10B981" },
                new ModuleInfo { Title = "Suppression d'Equipement",  Description = "Archiver ou supprimer des équipements obsolètes.",   IconKind = "Trash",  Color = "#EF4444" }
                new ModuleInfo { Title = "Suppression d'Equipement",  Description = "Archiver ou supprimer des équipements obsolètes.",   IconKind = "Trash",  Color = "#EF4444" }
                new ModuleInfo { Title = "Suppression d'Equipement",  Description = "Archiver ou supprimer des équipements obsolètes.",   IconKind = "Trash",  Color = "#EF4444" }
                new ModuleInfo { Title = "Suppression d'Equipement",  Description = "Archiver ou supprimer des équipements obsolètes.",   IconKind = "Trash",  Color = "#EF4444" }
                new ModuleInfo { Title = "Suppression d'Equipement",  Description = "Archiver ou supprimer des équipements obsolètes.",   IconKind = "Trash",  Color = "#EF4444" }
                new ModuleInfo { Title = "Suppression d'Equipement",  Description = "Archiver ou supprimer des équipements obsolètes.",   IconKind = "Trash",  Color = "#EF4444" }
            };
            
            _filteredModules = new ObservableCollection<ModuleInfo>(Modules);

            NavigateToModuleCommand = new RelayCommand(p => _mainViewModel.NavigateToModule((string)p!));
        }

        private void FilterModules()
        {
            if (string.IsNullOrWhiteSpace(SearchQuery))
            {
                FilteredModules = new ObservableCollection<ModuleInfo>(Modules);
            }
            else
            {
                var filtered = Modules.Where(m => m.Title.Contains(SearchQuery, System.StringComparison.OrdinalIgnoreCase)).ToList();
                FilteredModules = new ObservableCollection<ModuleInfo>(filtered);
            }
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
