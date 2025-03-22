using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.VisualBasic.FileIO; // For CSV parsing
using Microsoft.Win32;
using OfficeOpenXml; // For XLSX
using NPOI.SS.UserModel; // For XLS
using NPOI.XSSF.UserModel; // For XLSX (NPOI)
using NPOI.HSSF.UserModel; // For XLS (NPOI)
using System.Windows.Media.Imaging;

namespace ExcelCompareApp
{
    public partial class MainWindow : Window
    {
        // Variables to store file paths
        private string _excel1Path;
        private string _excel2Path;

        // List to store sheet mappings
        private List<SheetMapping> _sheetMappings = new List<SheetMapping>();

        public MainWindow()
        {
            InitializeComponent();
            this.Icon = new BitmapImage(new Uri("pack://application:,,,/app-icon.ico"));
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Required for EPPlus
        }

        // Class to represent sheet mappings
        public class SheetMapping
        {
            public string Excel1Sheet { get; set; }
            public int Excel1HeaderRow { get; set; }
            public string Excel2Sheet { get; set; }
            public int Excel2HeaderRow { get; set; }
            public List<string> Excel1Headers { get; set; }
            public List<string> Excel2Headers { get; set; }

            public ComboBox Excel1ColumnDropdown { get; set; } // Store dropdown for Excel 1
            public ComboBox Excel2ColumnDropdown { get; set; } // Store dropdown for Excel 2
        }

        // Upload Excel 1
        private void UploadExcel1_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.csv"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _excel1Path = openFileDialog.FileName;
                Excel1FilePath.Text = _excel1Path;
                LoadExcelSheets(_excel1Path, Excel1Sheets);
            }
        }

        // Upload Excel 2
        private void UploadExcel2_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.csv"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _excel2Path = openFileDialog.FileName;
                Excel2FilePath.Text = _excel2Path;
                LoadExcelSheets(_excel2Path, Excel2Sheets);
            }
        }

        // Load sheets from Excel file into ComboBox
        private void LoadExcelSheets(string filePath, ComboBox comboBox)
        {
            var fileExtension = Path.GetExtension(filePath).ToLower();

            if (fileExtension == ".xlsx" || fileExtension == ".xls")
            {
                // Use NPOI for XLSX and XLS files
                IWorkbook workbook = null;
                FileStream stream = null;

                try
                {
                    stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                    if (fileExtension == ".xlsx")
                    {
                        workbook = new XSSFWorkbook(stream); // XLSX
                    }
                    else
                    {
                        workbook = new HSSFWorkbook(stream); // XLS
                    }

                    var sheetNames = new List<string>();
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        sheetNames.Add(workbook.GetSheetName(i));
                    }

                    comboBox.ItemsSource = sheetNames;
                }
                catch (IOException ex)
                {
                    // Handle the case where the file is busy or open elsewhere
                    MessageBox.Show("The file is busy and might be open in another program. Please close the file and try again.", "File Busy", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                finally
                {
                    // Ensure the workbook and stream are properly disposed
                    workbook?.Close();
                    stream?.Close();
                }
            }
            else if (fileExtension == ".csv")
            {
                // For CSV files, treat the entire file as a single sheet
                comboBox.ItemsSource = new List<string> { "Sheet1" }; // Default sheet name for CSV
            }
            else
            {
                MessageBox.Show("Unsupported file format.");
            }
        }
        
        // Add sheet mapping
        private void AddSheetMapping_Click(object sender, RoutedEventArgs e)
        {
            if (Excel1Sheets.SelectedItem == null || Excel2Sheets.SelectedItem == null)
            {
                MessageBox.Show("Please select sheets from both Excel files.");
                return;
            }

            if (!int.TryParse(Excel1HeaderRow.Text, out int excel1HeaderRow) || excel1HeaderRow < 1)
            {
                MessageBox.Show("Please enter a valid header row index for Excel 1 (>= 1).");
                return;
            }

            if (!int.TryParse(Excel2HeaderRow.Text, out int excel2HeaderRow) || excel2HeaderRow < 1)
            {
                MessageBox.Show("Please enter a valid header row index for Excel 2 (>= 1).");
                return;
            }

            var mapping = new SheetMapping
            {
                Excel1Sheet = Excel1Sheets.SelectedItem.ToString(),
                Excel1HeaderRow = excel1HeaderRow,
                Excel2Sheet = Excel2Sheets.SelectedItem.ToString(),
                Excel2HeaderRow = excel2HeaderRow,
                Excel1Headers = LoadHeaders(_excel1Path, Excel1Sheets.SelectedItem.ToString(), excel1HeaderRow),
                Excel2Headers = LoadHeaders(_excel2Path, Excel2Sheets.SelectedItem.ToString(), excel2HeaderRow)
            };

            _sheetMappings.Add(mapping);
            RefreshMappingsList();
            AddColumnSelectionDropdowns(mapping);
        }

        // Load headers from a specific sheet and row
        private List<string> LoadHeaders(string filePath, string sheetName, int headerRow)
        {
            var fileExtension = Path.GetExtension(filePath).ToLower();

            if (fileExtension == ".xlsx" || fileExtension == ".xls")
            {
                // Use NPOI for XLSX and XLS files
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook;
                    if (fileExtension == ".xlsx")
                    {
                        workbook = new XSSFWorkbook(stream); // XLSX
                    }
                    else
                    {
                        workbook = new HSSFWorkbook(stream); // XLS
                    }

                    var sheet = workbook.GetSheet(sheetName);
                    var headerRowData = sheet.GetRow(headerRow - 1); // NPOI uses 0-based indexing

                    return headerRowData.Cells.Select(cell => cell.ToString()).ToList();
                }
            }
            else if (fileExtension == ".csv")
            {
                // Use TextFieldParser for CSV files
                using (var parser = new TextFieldParser(filePath))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");

                    // Read the header row
                    if (!parser.EndOfData)
                    {
                        return parser.ReadFields().ToList();
                    }
                }
            }

            return new List<string>();
        }

        // Add column selection dropdowns for a mapping
        private void AddColumnSelectionDropdowns(SheetMapping mapping)
        {
            var stackPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 5, 0, 5) };

            mapping.Excel1ColumnDropdown = new ComboBox
            {
                ItemsSource = mapping.Excel1Headers,
                Width = 150,
                Margin = new Thickness(0, 0, 10, 0)
            };

            mapping.Excel2ColumnDropdown = new ComboBox
            {
                ItemsSource = mapping.Excel2Headers,
                Width = 150,
                Margin = new Thickness(0, 0, 10, 0)
            };

            stackPanel.Children.Add(mapping.Excel1ColumnDropdown);
            stackPanel.Children.Add(mapping.Excel2ColumnDropdown);
            ColumnSelectionPanel.Children.Add(stackPanel);
        }

        // Remove a mapping
        private void RemoveMapping_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var mapping = button?.DataContext as SheetMapping;

            if (mapping != null)
            {
                _sheetMappings.Remove(mapping);
                RefreshMappingsList();
                RefreshColumnSelectionDropdowns();
            }
        }

        // Refresh the mappings list
        private void RefreshMappingsList()
        {
            SheetMappings.ItemsSource = null;
            SheetMappings.ItemsSource = _sheetMappings;
        }

        // Refresh the column selection dropdowns
        private void RefreshColumnSelectionDropdowns()
        {
            ColumnSelectionPanel.Children.Clear();
            foreach (var mapping in _sheetMappings)
            {
                AddColumnSelectionDropdowns(mapping);
            }
        }

        // Compare and download results
        private void CompareAndDownload_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_excel1Path) || string.IsNullOrEmpty(_excel2Path))
            {
                MessageBox.Show("Please upload both Excel files.");
                return;
            }

            if (_sheetMappings.Count == 0)
            {
                MessageBox.Show("Please add at least one sheet mapping.");
                return;
            }

            try
            {
                foreach (var mapping in _sheetMappings)
                {
                    // Get selected columns from dropdowns
                    var excel1Column = mapping.Excel1ColumnDropdown.SelectedItem?.ToString();
                    var excel2Column = mapping.Excel2ColumnDropdown.SelectedItem?.ToString();

                    if (string.IsNullOrEmpty(excel1Column) || string.IsNullOrEmpty(excel2Column))
                    {
                        MessageBox.Show("Please select columns to compare for all mappings.");
                        return;
                    }

                    var excel1Data = LoadColumnData(_excel1Path, mapping.Excel1Sheet, mapping.Excel1HeaderRow, excel1Column);
                    var excel2Data = LoadColumnData(_excel2Path, mapping.Excel2Sheet, mapping.Excel2HeaderRow, excel2Column);

                    // Find differences
                    var differences = excel1Data.Except(excel2Data).ToList();

                    // Save differences to a new Excel file
                    var outputPackage = new ExcelPackage();
                    var outputSheet = outputPackage.Workbook.Worksheets.Add($"{mapping.Excel1Sheet}_vs_{mapping.Excel2Sheet}");

                    // Add headers
                    outputSheet.Cells[1, 1].Value = $"{excel1Column} (Not in {excel2Column})";

                    // Add differences
                    for (int i = 0; i < differences.Count; i++)
                    {
                        outputSheet.Cells[i + 2, 1].Value = differences[i];
                    }

                    // Save the file
                    var saveFileDialog = new SaveFileDialog
                    {
                        Filter = "Excel Files|*.xlsx",
                        FileName = $"{mapping.Excel1Sheet}_vs_{mapping.Excel2Sheet}_Comparison_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx"
                    };

                    if (saveFileDialog.ShowDialog() == true)
                    {
                        outputPackage.SaveAs(new FileInfo(saveFileDialog.FileName));
                        MessageBox.Show($"Comparison for {mapping.Excel1Sheet} vs {mapping.Excel2Sheet} saved successfully!");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

        // Load column data from a specific sheet and column
        private List<string> LoadColumnData(string filePath, string sheetName, int headerRow, string columnName)
        {
            var fileExtension = Path.GetExtension(filePath).ToLower();

            if (fileExtension == ".xlsx" || fileExtension == ".xls")
            {
                // Use NPOI for XLSX and XLS files
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook;
                    if (fileExtension == ".xlsx")
                    {
                        workbook = new XSSFWorkbook(stream); // XLSX
                    }
                    else
                    {
                        workbook = new HSSFWorkbook(stream); // XLS
                    }

                    var sheet = workbook.GetSheet(sheetName);
                    var headerRowData = sheet.GetRow(headerRow - 1); // NPOI uses 0-based indexing

                    // Find the column index
                    int columnIndex = -1;
                    for (int i = 0; i < headerRowData.Cells.Count; i++)
                    {
                        if (headerRowData.Cells[i].ToString() == columnName)
                        {
                            columnIndex = i;
                            break;
                        }
                    }

                    if (columnIndex == -1)
                    {
                        throw new Exception($"Column '{columnName}' not found in the sheet.");
                    }

                    // Read the column data
                    var columnData = new List<string>();
                    for (int i = headerRow; i <= sheet.LastRowNum; i++)
                    {
                        var row = sheet.GetRow(i);
                        if (row != null && row.Cells.Count > columnIndex)
                        {
                            columnData.Add(row.Cells[columnIndex].ToString());
                        }
                    }

                    return columnData;
                }
            }
            else if (fileExtension == ".csv")
            {
                // Use TextFieldParser for CSV files
                using (var parser = new TextFieldParser(filePath))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");

                    // Read the header row
                    if (!parser.EndOfData)
                    {
                        var headers = parser.ReadFields();
                        int columnIndex = Array.IndexOf(headers, columnName);

                        if (columnIndex == -1)
                        {
                            throw new Exception($"Column '{columnName}' not found in the CSV file.");
                        }

                        // Read the column data
                        var columnData = new List<string>();
                        while (!parser.EndOfData)
                        {
                            var fields = parser.ReadFields();
                            if (fields.Length > columnIndex)
                            {
                                columnData.Add(fields[columnIndex]);
                            }
                        }

                        return columnData;
                    }
                }
            }

            throw new Exception("Unsupported file format.");
        }
    }
}