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
                new ModuleInfo { Number="01", Title = "Création de Postes Techniques", Description = "Créer en masse de nouveaux postes techniques dans SAP via Excel.", IconKind = "\xE710", Color = "#3B82F6", Version = "v1.0.0", HealthStatus = "Optimal", Status="ACTIVE", StatusForegroundColor="#10B981", StatusBackgroundColor="#D1FAE5" },
                new ModuleInfo { Number="02", Title = "Modification de Postes Techniques", Description = "Modifier en masse les postes techniques existants dans SAP via Excel.", IconKind = "\xE70F", Color = "#10B981", Version = "v1.0.0", HealthStatus = "Optimal", Status="UPDATING", StatusForegroundColor="#F59E0B", StatusBackgroundColor="#FEF3C7" },
                new ModuleInfo { Number="03", Title = "Suppression de Postes Techniques",  Description = "Supprimer en masse les postes techniques dans SAP via Excel.", IconKind = "\xE74D", Color = "#EF4444", Version = "v1.0.0", HealthStatus = "Optimal", Status="ERROR", StatusForegroundColor="#EF4444", StatusBackgroundColor="#FEE2E2" },
                new ModuleInfo { Number="04", Title = "Création d'Equipements", Description = "Créer en masse de nouveaux équipements dans SAP via Excel.", IconKind = "\xE710", Color = "#3B82F6", Version = "v1.0.0", HealthStatus = "Optimal", Status="ACTIVE", StatusForegroundColor="#10B981", StatusBackgroundColor="#D1FAE5" },
                new ModuleInfo { Number="05", Title = "Modification d'Equipements", Description = "Modifier en masse les équipements existants dans SAP via Excel.", IconKind = "\xE70F", Color = "#10B981", Version = "v1.0.0", HealthStatus = "Optimal", Status="ACTIVE", StatusForegroundColor="#10B981", StatusBackgroundColor="#D1FAE5" },
                new ModuleInfo { Number="06", Title = "Suppression d'Equipements",  Description = "Supprimer en masse des équipements dans SAP via Excel.", IconKind = "\xE74D", Color = "#EF4444", Version = "v1.0.0", HealthStatus = "Optimal", Status="ACTIVE", StatusForegroundColor="#10B981", StatusBackgroundColor="#D1FAE5" },
                new ModuleInfo { Number="07", Title = "Création d'Equipements", Description = "Créer en masse de nouveaux équipements dans SAP via Excel.", IconKind = "\xE710", Color = "#3B82F6", Version = "v1.0.0", HealthStatus = "Optimal", Status="ACTIVE", StatusForegroundColor="#10B981", StatusBackgroundColor="#D1FAE5" },
                new ModuleInfo { Number="08", Title = "Modification d'Equipements", Description = "Modifier en masse les équipements existants dans SAP via Excel.", IconKind = "\xE70F", Color = "#10B981", Version = "v1.0.0", HealthStatus = "Optimal", Status="ACTIVE", StatusForegroundColor="#10B981", StatusBackgroundColor="#D1FAE5" },
                new ModuleInfo { Number="09", Title = "Suppression d'Equipements",  Description = "Supprimer en masse des équipements dans SAP via Excel.", IconKind = "\xE74D", Color = "#EF4444", Version = "v1.0.0", HealthStatus = "Optimal", Status="ACTIVE", StatusForegroundColor="#10B981", StatusBackgroundColor="#D1FAE5" }
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
        public string Number { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string IconKind { get; set; } = string.Empty;
        public string Color { get; set; } = string.Empty;
        public string Version { get; set; } = string.Empty;
        public string HealthStatus { get; set; } = string.Empty;
        public string Status { get; set; } = "ACTIVE";
        public string StatusForegroundColor { get; set; } = "#10B981"; // AccentGreen
        public string StatusBackgroundColor { get; set; } = "#D1FAE5"; // GreenPastel
    }
}
