using Avalonia.Controls;
using Avalonia.ReactiveUI;
using ReactiveUI;
using SQLScript2XLSX.ViewModels;
using System;
using System.Reactive;
using System.Threading.Tasks;

namespace SQLScript2XLSX.Views
{
    public partial class MainWindow : ReactiveWindow<MainWindowViewModel>
    {
        public MainWindow()
        {
            InitializeComponent();
            this.WhenActivated(d => d(ViewModel!.SaveFileInteraction.RegisterHandler(DoShowDialogAsync)));

        }
        private async Task DoShowDialogAsync(InteractionContext<Unit, string?> interaction)
        {
            var dialog = new SaveFileDialog();
            dialog.DefaultExtension = ".xlsx";
            dialog.Filters?.Add(new FileDialogFilter() { Name = "Excel Workbook", Extensions = { "xlsx" } });
            dialog.Filters?.Add(new FileDialogFilter() { Name = "All Files", Extensions = { "*" } });
            dialog.Directory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var result = await dialog.ShowAsync(this);
            interaction.SetOutput(result);
        }
    }
}
