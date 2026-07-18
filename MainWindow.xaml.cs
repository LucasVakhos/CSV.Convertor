using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;

namespace CSV.Convertor
{
    public partial class MainWindow : Window
    {
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

            try
            {
                btnLoad.IsEnabled = false;
                progressBar.Value = 0;
                lblProgress.Text = "0%";
                lblStatus.Text = "Конвертация...";

                string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string outputDirectory = Path.Combine(exeDirectory, "output");
                Directory.CreateDirectory(outputDirectory);

                string originalFileName = Path.GetFileNameWithoutExtension(openDlg.FileName);
                string xlsxPath = Path.Combine(outputDirectory, $"{originalFileName}.xlsx");

                var progress = new Progress<int>(p =>
                {
                    progressBar.Value = p;
                    lblProgress.Text = $"{p}%";
                });

                await Task.Run(() => CsvConverter.Convert(openDlg.FileName, xlsxPath, progress));

                lblStatus.Text = $"✅ Готово! Файл сохранён: {Path.GetFileName(xlsxPath)}";
                MessageBox.Show($"Конвертация завершена!\nФайл сохранён:\n{xlsxPath}", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);

                OpenFile(xlsxPath);
            }
            catch (Exception ex)
            {
                lblStatus.Text = $"❌ Ошибка: {ex.Message}";
                MessageBox.Show($"Ошибка конвертации:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnLoad.IsEnabled = true;
            }
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
    }
}