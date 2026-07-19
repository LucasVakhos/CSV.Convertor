using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace CSV.Convertor
{
    public partial class MainWindow : Window
    {
        private string? lastConvertedFile = null;

        public MainWindow()
        {
            InitializeComponent();
        }

        private async void BtnLoad_Click(object sender, RoutedEventArgs e)
        {
            var openDlg = new OpenFileDialog
            {
                Filter = "CSV файлы (*.csv)|*.csv",
                Title = "Выберите CSV файл для конвертации"
            };

            if (openDlg.ShowDialog() != true) return;

            await ConvertFile(openDlg.FileName);
        }

        private async Task ConvertFile(string csvPath)
        {
            try
            {
                btnLoad.IsEnabled = false;
                btnOpenXlsx.IsEnabled = false;
                btnExit.IsEnabled = false;

                progressBar.Value = 0;
                lblProgress.Text = "0%";
                lblStatus.Text = "Конвертация...";

                // Получаем разделитель из интерфейса
                char? delimiter = GetSelectedDelimiter();

                // Используем папку «Документы» пользователя, чтобы избежать проблем
                // с правами записи при установке приложения в Program Files
                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string outputDirectory = Path.Combine(documentsPath, "CSV.Convertor", "output");
                Directory.CreateDirectory(outputDirectory);

                string originalFileName = Path.GetFileNameWithoutExtension(csvPath);
                string xlsxPath = Path.Combine(outputDirectory, $"{originalFileName}.xlsx");

                var progress = new Progress<int>(p =>
                {
                    progressBar.Value = p;
                    lblProgress.Text = $"{p}%";
                });

                await Task.Run(() => CsvConverter.Convert(csvPath, xlsxPath, delimiter, progress!));

                lastConvertedFile = xlsxPath;

                lblStatus.Text = $"✅ Готово! Файл сохранён: {Path.GetFileName(xlsxPath)}";

                var result = MessageBox.Show(
                    $"Конвертация завершена!\nФайл сохранён:\n{xlsxPath}\n\nХотите конвертировать ещё один файл?",
                    "Успех",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    OpenFile(xlsxPath);

                    var openDlg = new OpenFileDialog
                    {
                        Filter = "CSV файлы (*.csv)|*.csv",
                        Title = "Выберите следующий CSV файл для конвертации"
                    };

                    if (openDlg.ShowDialog() == true)
                    {
                        await ConvertFile(openDlg.FileName);
                    }
                    else
                    {
                        Environment.Exit(0);
                    }
                }
                else
                {
                    OpenFile(xlsxPath);
                    Environment.Exit(0);
                }
            }
            catch (Exception ex)
            {
                lblStatus.Text = $"❌ Ошибка: {ex.Message}";
                MessageBox.Show($"Ошибка конвертации:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnLoad.IsEnabled = true;
                btnOpenXlsx.IsEnabled = true;
                btnExit.IsEnabled = true;
            }
        }

        private char? GetSelectedDelimiter()
        {
            if (chkAutoDetect.IsChecked == true)
                return null;

            if (cmbDelimiter.SelectedItem is ComboBoxItem item && item.Tag is string tag)
            {
                if (tag == "\\t")
                    return '\t';
                return tag[0];
            }

            return ';';
        }

        private void BtnOpenXlsx_Click(object sender, RoutedEventArgs e)
        {
            var openDlg = new OpenFileDialog
            {
                Filter = "Excel файлы (*.xlsx)|*.xlsx",
                Title = "Выберите Excel файл для открытия"
            };

            if (openDlg.ShowDialog() == true)
            {
                OpenFile(openDlg.FileName);
            }
        }

        private void OpenFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = filePath,
                        UseShellExecute = true
                    };
                    Process.Start(psi);
                    lblStatus.Text = $"📂 Файл открыт: {Path.GetFileName(filePath)}";
                }
                else
                {
                    MessageBox.Show("Файл не найден.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось открыть файл:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            Environment.Exit(0);
        }
    }
}