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

        public void NavigateToModule(string moduleTitle)
        {
            CurrentViewModel = new ModuleDetailViewModel(this, moduleTitle);
        }
    }
}
