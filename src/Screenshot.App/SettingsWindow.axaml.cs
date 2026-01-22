using Avalonia.Controls;
using Screenshot.App.ViewModels;

namespace Screenshot.App
{
    public partial class SettingsWindow : Window
    {
        public SettingsWindow()
        {
            InitializeComponent();
        }

        private async void OnExportSettingsClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel vm) return;
            var dialog = new SaveFileDialog
            {
                Title = "导出设置",
                InitialFileName = "ScreenshotV4.settings.json"
            };
            dialog.Filters?.Add(new FileDialogFilter { Name = "JSON", Extensions = { "json" } });
            var path = await dialog.ShowAsync(this);
            if (!string.IsNullOrWhiteSpace(path))
            {
                vm.ExportSettingsTo(path);
            }
        }

        private async void OnImportSettingsClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel vm) return;
            var dialog = new OpenFileDialog
            {
                Title = "导入设置",
                AllowMultiple = false
            };
            dialog.Filters?.Add(new FileDialogFilter { Name = "JSON", Extensions = { "json" } });
            var paths = await dialog.ShowAsync(this);
            if (paths != null && paths.Length > 0)
            {
                vm.ImportSettingsFrom(paths[0]);
            }
        }
    }
}
