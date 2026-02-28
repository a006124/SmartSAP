using SmartSAP.ViewModels.Modules;

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

        public MainViewModel()
        {
            _currentViewModel = new LibraryViewModel(this);
        }

        public void NavigateToLibrary()
        {
            CurrentViewModel = new LibraryViewModel(this);
        }

        public void NavigateToModule(string moduleNumber)
        {
            switch (moduleNumber)
            {
                case "01":
                    CurrentViewModel = new Module01ViewModel(this, "Création de Postes Techniques");
                    break;
                // D'autres modules viendront s'ajouter ici par la suite
                default:
                    // Pour l'instant, si on clique sur un autre module non implémenté, on peut instancier un ViewModel basique ou ne rien faire
                    // CurrentViewModel = new NotImplementedViewModel();
                    break;
            }
        }
    }
}
