using BindDataFromDiffSource.Models;
using Syncfusion.XlsIO;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text.Json;

namespace BindDataFromDiffSource.ViewModels
{
    /// <summary>
    /// ViewModel demonstrating data binding from JSON, CSV, and Excel sources
    /// using only built-in .NET libraries (no third-party packages)
    /// </summary>
    public class MainViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        // Collections for binding to SparkCharts
        public ObservableCollection<DataPoint> JsonData { get; set; } = new();
        public ObservableCollection<DataPoint> CsvData { get; set; } = new();
        public ObservableCollection<DataPoint> ExcelData { get; set; } = new();

        public MainViewModel()
        {
            LoadDataAsync();
        }

        /// <summary>
        /// Load data from all three sources (JSON, CSV, Excel)
        /// </summary>
        private async void LoadDataAsync()
        {
            await LoadJsonDataAsync();
            await LoadCsvDataAsync();
            await LoadExcelDataAsync();
        }

        #region Load JSON Data
        /// <summary>
        /// Load data from JSON file in Resources/Raw
        /// </summary>
        private async Task LoadJsonDataAsync()
        {
            try
            {
                // Read JSON file from Resources/Raw
                using var stream = await FileSystem.OpenAppPackageFileAsync("SalesData.json");
                using var reader = new StreamReader(stream);
                var jsonContent = await reader.ReadToEndAsync();

                // Deserialize JSON - the file has "Sales" property, not "Value"
                var jsonDocument = JsonDocument.Parse(jsonContent);
                
                JsonData.Clear();
                foreach (var element in jsonDocument.RootElement.EnumerateArray())
                {
                    var month = element.GetProperty("Month").GetString() ?? "";
                    var sales = element.GetProperty("Sales").GetDouble();
                    
                    JsonData.Add(new DataPoint
                    {
                        Month = month,
                        Value = sales
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading JSON: {ex.Message}");
            }
        }
        #endregion

        #region Load CSV Data
        /// <summary>
        /// Load data from CSV file in Resources/Raw using built-in .NET StreamReader
        /// </summary>
        private async Task LoadCsvDataAsync()
        {
            try
            {
                // Read CSV file from Resources/Raw
                using var stream = await FileSystem.OpenAppPackageFileAsync("RevenueData.csv");
                using var reader = new StreamReader(stream);

                // Read and skip header line
                var headerLine = await reader.ReadLineAsync();

                // Read CSV records
                CsvData.Clear();
                string? line;
                while ((line = await reader.ReadLineAsync()) != null)
                {
                    var parts = line.Split(',');
                    if (parts.Length >= 2)
                    {
                        CsvData.Add(new DataPoint
                        {
                            Month = parts[0].Trim(),
                            Value = double.Parse(parts[1].Trim(), CultureInfo.InvariantCulture)
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading CSV: {ex.Message}");
            }
        }
        #endregion

        #region Load Excel Data
        /// <summary>
        /// Load data from Excel (.xlsx) file using Syncfusion.XlsIO
        /// </summary>
        private async Task LoadExcelDataAsync()
        {
            try
            {
                // Read Excel file from Resources/Raw
                using var stream = await FileSystem.OpenAppPackageFileAsync("ExpensesData.xlsx");
                
                // Create Excel engine and open workbook
                using var excelEngine = new ExcelEngine();
                var application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                
                var workbook = application.Workbooks.Open(stream);
                var worksheet = workbook.Worksheets[0]; // First sheet

                ExcelData.Clear();

                // Read data starting from row 2 (skip header)
                for (int row = 2; row <= worksheet.UsedRange.LastRow; row++)
                {
                    var month = worksheet.Range[row, 1].Text; // Column A (text)
                    var expenseCell = worksheet.Range[row, 2]; // Column B (number)

                    if (!string.IsNullOrEmpty(month) && expenseCell.Value != null)
                    {
                        // Use Value property for numeric data instead of Text
                        double expense = Convert.ToDouble(expenseCell.Value);
                        
                        ExcelData.Add(new DataPoint
                        {
                            Month = month,
                            Value = expense
                        });
                    }
                }

                workbook.Close();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading Excel: {ex.Message}");
            }
        }
        #endregion

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
