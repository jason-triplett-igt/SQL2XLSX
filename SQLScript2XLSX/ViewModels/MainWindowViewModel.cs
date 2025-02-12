using Avalonia;
using Avalonia.Controls;
using ReactiveUI;
using ReactiveUI.Fody.Helpers;
using SQLScript2XLSX.Models;
using System;
using System.Collections.Generic;
using System.Reactive;
using System.Reactive.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SQLScript2XLSX.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    {
        public MainWindowViewModel()
        {
            SaveFileInteraction = new Interaction<Unit, string?>();
            SaveFileCommand = ReactiveCommand.CreateFromTask<string?>(async () =>
            {
                string? savepath = await SaveFileInteraction.Handle(Unit.Default);
                return savepath;
            });
            ExportDataCommand = ReactiveCommand.CreateFromTask( async () =>
            {
                ExceptionMessage = "";
                CancellationTokenSource = new CancellationTokenSource();
                await ExportDataToXLSXfromSQL.ExportAsync(this, CancellationTokenSource.Token);
                ExceptionMessage = "Export Complete!";
            });
            CopyExceptionMessageToClipboardCommand = ReactiveCommand.CreateFromTask (async () =>
            {
                if(Application.Current is not null && Application.Current.Clipboard is not null)
                {
                    await Application.Current.Clipboard.SetTextAsync(ExceptionMessage);
                }
            });
            CancelExport = ReactiveCommand.Create(() =>
            {
                CancellationTokenSource?.Cancel();
            });
    
            this.WhenAnyObservable(x => x.SaveFileCommand)
                .Subscribe(x => OutputPath = x);
            ExportDataCommand.ThrownExceptions.Subscribe(error => ExceptionMessage = error.Message);
            ExceptionMessage = "";
            
        }

        [Reactive]
        public string? Script { get; set; }
        [Reactive]
        public string? Name { get; set; }
        [Reactive]
        public string? Datasource { get; set; }
        [Reactive]
        public string? InitialCatalog { get; set; }
        [Reactive]
        public bool IntegratedSecurity { get; set; }
        [Reactive]
        public string? Username { get; set; }
        [Reactive]
        public string? Password { get; set; }
        [Reactive]
        public string? OutputPath { get; set; }
        [Reactive]
        public string? ExceptionMessage { get; set; }
        public ReactiveCommand<Unit, string?> SaveFileCommand { get; }
        public Interaction<Unit, string?> SaveFileInteraction { get; }
        public ReactiveCommand<Unit, Unit> ExportDataCommand { get; set; }
        public ReactiveCommand<Unit, Unit> CopyExceptionMessageToClipboardCommand { get; set; }
        [Reactive]
        public CancellationTokenSource? CancellationTokenSource { get; set; }
        public ReactiveCommand<Unit, Unit> CancelExport { get; set; }
    }
}
