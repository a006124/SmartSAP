using System.Linq;
using System.Windows;
using SmartSAP.ViewModels.Modules;

namespace SmartSAP.Views
{
    public partial class ModuleDetailView : System.Windows.Controls.UserControl
    {
        public ModuleDetailView()
        {
            InitializeComponent();
        }

        private void LogSection_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                {
                    string droppedFile = files.FirstOrDefault(f => f.EndsWith(".xlsx", System.StringComparison.OrdinalIgnoreCase) || f.EndsWith(".xls", System.StringComparison.OrdinalIgnoreCase));
                    
                    if (!string.IsNullOrEmpty(droppedFile))
                    {
                        var viewModel = this.DataContext as ModuleDetailViewModelBase;
                        if (viewModel != null)
                        {
                            viewModel.HandleDroppedExcelFile(droppedFile);
                        }
                    }
                }
            }
        }
    }
}
