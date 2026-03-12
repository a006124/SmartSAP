using System.Windows.Controls;

namespace SmartSAP.Views.Modules
{
    public partial class Module02View : UserControl
    {
        public Module02View()
        {
            InitializeComponent();
        }

        private void LogSection_DragEnter(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                e.Effects = System.Windows.DragDropEffects.Copy;
            }
            else
            {
                e.Effects = System.Windows.DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void LogSection_Drop(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                {
                    string droppedFile = System.Linq.Enumerable.FirstOrDefault(files, f =>
                        f.EndsWith(".xlsx", System.StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".xls", System.StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".txt", System.StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".csv", System.StringComparison.OrdinalIgnoreCase));

                    if (!string.IsNullOrEmpty(droppedFile))
                    {
                        var viewModel = this.DataContext as SmartSAP.ViewModels.Modules.ModuleDetailViewModelBase;
                        if (viewModel != null)
                        {
                            viewModel.HandleDroppedFile(droppedFile);
                        }
                    }
                }
            }
            e.Handled = true;
        }

    }
}
