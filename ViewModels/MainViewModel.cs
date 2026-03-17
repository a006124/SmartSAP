using SmartSAP.ViewModels.Modules;
using System.Windows.Media;

namespace SmartSAP.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private ViewModelBase _currentViewModel;
        public ViewModelBase CurrentViewModel
        {
            get => _currentViewModel;
            set => SetProperty(ref _currentViewModel, value);
        }

        private bool _isSAPConnected;
        public bool IsSAPConnected
        {
            get => _isSAPConnected;
            set 
            {
                if (SetProperty(ref _isSAPConnected, value))
                {
                    OnPropertyChanged(nameof(SAPStatusBrush));
                }
            }
        }

        private string _sapInstanceInfo = "Non connecté";
        public string SAPInstanceInfo
        {
            get => _sapInstanceInfo;
            set => SetProperty(ref _sapInstanceInfo, value);
        }

        public Brush SAPStatusBrush => IsSAPConnected ? new SolidColorBrush(Color.FromRgb(16, 185, 129)) : new SolidColorBrush(Color.FromRgb(239, 68, 68));

        public MainViewModel()
        {
            _currentViewModel = new LibraryViewModel(this);
        }

        public void NavigateToLibrary()
        {
            CurrentViewModel = new LibraryViewModel(this);
        }

        public void NavigateToModule(string parameter)
        {
            if (string.IsNullOrEmpty(parameter)) return;

            string key = parameter.Trim();

            // Supporte à la fois le numéro (01) et le titre pour plus de robustesse
            switch (key)
            {
                case "01":
                case "Création de Postes Techniques":
                    CurrentViewModel = new Module01ViewModel(this, "Création de Postes Techniques");
                    break;
                case "02":
                case "Modification de Postes Techniques":
                    CurrentViewModel = new Module02ViewModel(this, "Modification de Postes Techniques");
                    break;
                case "03":
                case "Suppression de Postes Techniques":
                    CurrentViewModel = new Module03ViewModel(this, "Suppression de Postes Techniques");
                    break;
                case "04":
                case "Création d'Equipements":
                    CurrentViewModel = new Module04ViewModel(this, "Création d'Equipements");
                    break;
                case "05":
                case "Modification d'Equipements":
                    CurrentViewModel = new Module05ViewModel(this, "Modification d'Equipements");
                    break;
                case "06":
                case "Suppression d'Equipements":
                    CurrentViewModel = new Module06ViewModel(this, "Suppression d'Equipements");
                    break;
                case "07":
                case "Extraction de Gammes":
                    CurrentViewModel = new Module07ViewModel(this, "Extraction de Gammes");
                    break;
                case "08":
                    CurrentViewModel = new Module08ViewModel(this, "Modification d'Equipements");
                    break;
                case "09":
                    CurrentViewModel = new Module09ViewModel(this, "Suppression d'Equipements");
                    break;
            }
        }
    }
}
