using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO; // Required for file operations
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private string selectedFilePath; // Store the selected Excel file path
        private Timer timer; // Timer to read the last line every 10 seconds
        // Member of type Dictionary<string, string> initialized using InitializeDictionary
        private Dictionary<string, string> writeDictionary;
        private struct GasData
        {
            public double SO2;
            public double NO_NO2;
            public double CO;
            public double CO2;
            public double O2;
            public double SPM;
            public double Temp;
        }
        private GasData gasData; // Declare gasData as a class-level member
        public Form1()
        {
            InitializeComponent();
            InitializeCustomComponents();
            writeDictionary = InitializeDictionary(); // Initialize the dictionary
            this.FormClosing += Form1_FormClosing;
            GasData gasData = new GasData
            {
                SO2 = 0.0,
                NO_NO2 = 0.0,
                CO = 0.0,
                CO2 = 0.0,
                O2 = 0.0,
                SPM = 0.0,
                Temp = 0.0
            };
        }

        private void InitializeCustomComponents()
        {
            // Create a button for input file
            Button browseButton = new Button
            {
                Text = "Input File",
                Location = new System.Drawing.Point(10, 10), // Set the position
                Size = new System.Drawing.Size(100, 30) // Set the size
            };

            // Add click event handler for input file button
            browseButton.Click += BrowseButton_Click;

            // Add the button to the form
            this.Controls.Add(browseButton);

            // Create a button for creating target file
            Button createFileButton = new Button
            {
                Text = "Create Target File",
                Location = new System.Drawing.Point(10, 50), // Set the position below the first button
                Size = new System.Drawing.Size(150, 30) // Set the size
            };

            // Add click event handler for create file button
            createFileButton.Click += CreateFileButton_Click;

            // Add the button to the form
            this.Controls.Add(createFileButton);

            // Create a button for processing the file
            Button processFileButton = new Button
            {
                Text = "Process File",
                Location = new System.Drawing.Point(10, 90), // Set the position below the second button
                Size = new System.Drawing.Size(120, 30) // Set the size
            };

            // Add click event handler for process file button
            processFileButton.Click += ProcessFileButton_Click;

            // Add the button to the form
            this.Controls.Add(processFileButton);

            // Initialize the timer
            timer = new Timer
            {
                Interval = 60000 // Set interval to 60 seconds (60000 ms)
            };
            timer.Tick += Timer_Tick;
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // Set file filter to accept Excel files
                openFileDialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx";
                openFileDialog.Title = "Select an Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = openFileDialog.FileName; // Store the selected file path
                    MessageBox.Show($"Selected File: {selectedFilePath}", "File Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Start the timer after selecting the file
                    timer.Start();
                }
            }
        }

        private string targetFilePath; // Store the target file path

        private void CreateFileButton_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                // Set file filter to save as .csv file
                saveFileDialog.Filter = "CSV Files (*.csv)|*.csv";
                saveFileDialog.Title = "Create Target CSV File";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    targetFilePath = saveFileDialog.FileName;

                    try
                    {
                        // Create an empty .csv file
                        File.WriteAllText(targetFilePath, string.Empty);

                        MessageBox.Show($"Target CSV File Created: {targetFilePath}", "File Created", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error creating CSV file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ProcessFileButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFilePath))
            {
                MessageBox.Show("Please select an input Excel file first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            ReadLastLineOfExcel();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            ReadLastLineOfExcel();
            if (!string.IsNullOrEmpty(targetFilePath))
            {
                WriteDictionaryToFile(targetFilePath);
                ConvertCsvToExcel(targetFilePath); // Convert CSV to Excel after writing
            }
        }
        private Dictionary<string, string> InitializeDictionary()
        {
            // Create and initialize the dictionary
            Dictionary<string, string> dataDictionary = new Dictionary<string, string>
            {
                { "heading1", "    Green Power              " },
                { "heading2", "Parameters ,Unit ,Real Time  " },
                { "SO2",      "SO2         ,PPM,   0.0      " },
                { "NO+NO2",   "NO+NO2      ,PPM,   0.0      " },
                { "CO",       "CO          ,PPM,   0.0      " },
                { "O2",       "O2          ,  %,   0.0      " },
                { "SPM",      "SPM      ,mg/ml3,   0.0      " },
                { "Temp",     "Temp     ,  DegC,   0.0      " },
                { "Flow",     "Flow     , m3/hr,   0.0      " },
                { "DateTime", "                             " },
                { "heading3", "Supplied By                  " },
                { "heading4", "Enthalpy Asia Co Ltd         " }


            };

            return dataDictionary;
        }
        private void WriteDictionaryToFile(string filePath)
        {
            try
            {
                // Prepare the dictionary content for writing
                List<string> lines = new List<string>();
                foreach (var kvp in writeDictionary)
                {
                    lines.Add(kvp.Value);
                }

                // Write the dictionary content to the file
                File.WriteAllLines(filePath, lines);
            }
            catch (Exception ex)
            {
                // Handle any exceptions that occur during file writing
                MessageBox.Show($"Error writing dictionary to file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private string[] cellValues; // Declare cellValues as a class-level member
        private void ReadLastLineOfExcel()
        {
            if (string.IsNullOrEmpty(selectedFilePath))
            {
                MessageBox.Show("No file selected to process.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // Create an Excel application instance
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = null;
                Excel.Worksheet worksheet = null;

                try
                {
                    // Open the workbook in read-only mode
                    workbook = excelApp.Workbooks.Open(selectedFilePath, ReadOnly: true,Notify:true);
                    worksheet = workbook.Sheets[1]; // Get the first worksheet

                    // Get the last used row and column
                    //int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    Excel.Range lastCell = worksheet.Cells[worksheet.Rows.Count, 1].End(Excel.XlDirection.xlUp);
                    int lastRow = lastCell.Row;
                    Marshal.ReleaseComObject(lastCell);
                    Excel.Range lastCell1 = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                    int lastColumn = lastCell1.Column;
                    Marshal.ReleaseComObject(lastCell1);


                    // Array to store cell values
                    cellValues = new string[lastColumn];

                    for (int col = 1; col <= lastColumn; col++)
                    {
                        cellValues[col - 1] = worksheet.Cells[lastRow, col].Text.ToString(); // Store cell value in array
                    }

                    // Update the DateTime key in writeDictionary
                    if (writeDictionary.ContainsKey("DateTime"))
                    {
                        writeDictionary["DateTime"] = cellValues[0]; // Update DateTime key with cellValues[0]
                    }

                    writeforSO2();
                    writeforNOX();
                    writeforCO();
                    //writeforCO2();
                    writeforO2();
                    writeforSPM();
                    writeforTemp();
                    writeforFlow();

                    // Display cell values
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("Last Line of Excel File (Cell Values):");
                    foreach (var value in cellValues)
                    {
                        sb.AppendLine(value);
                    }

                    //MessageBox.Show(sb.ToString(), "Last Line Read", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally
                {
                    // Release COM objects
                    if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                    if (workbook != null)
                    {
                        workbook.Close(false);
                        Marshal.ReleaseComObject(workbook);
                    }
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private int position = 17;
        private void writeforSO2()
        {
            if (writeDictionary.ContainsKey("SO2"))
            {
                
                string currentValue = writeDictionary["SO2"];

                // Attempt to parse and round cellValues[1] and cellValues[2]
                string roundedValue1 = double.TryParse(cellValues[1], out double numericValue1)
                    ? Math.Round(numericValue1, 2).ToString().PadLeft(6)
                    : "  0.0"; // Default value if parsing fails
                gasData.SO2 = Math.Max(numericValue1, gasData.SO2);
                string roundedValue2 = Math.Round(gasData.SO2, 2).ToString().PadLeft(10);

                // Construct the updated value
                string updatedValue = currentValue.Substring(0, position) + roundedValue1;

                // Update the dictionary with the new value
                writeDictionary["SO2"] = updatedValue;
            }
        }
        private void writeforNOX()
        {
            if (writeDictionary.ContainsKey("NO+NO2"))
            {
                string currentValue = writeDictionary["NO+NO2"];

                // Attempt to parse and round cellValues[1] and cellValues[2]
                string roundedValue1 = double.TryParse(cellValues[2], out double numericValue1)
                    ? Math.Round(numericValue1, 2).ToString().PadLeft(6)
                    : "  0.0"; // Default value if parsing fails
                gasData.NO_NO2 = Math.Max(numericValue1, gasData.NO_NO2);
                string roundedValue2 = Math.Round(gasData.SO2, 2).ToString().PadLeft(10);

                // Construct the updated value
                string updatedValue = currentValue.Substring(0, position) + roundedValue1;

                // Update the dictionary with the new value
                writeDictionary["NO+NO2"] = updatedValue;
            }
        }
        private void writeforCO()
        {
            if (writeDictionary.ContainsKey("CO"))
            {
                string currentValue = writeDictionary["CO"];

                // Attempt to parse and round cellValues[3] and cellValues[4]
                string roundedValue1 = double.TryParse(cellValues[3], out double numericValue1)
                    ? Math.Round(numericValue1, 2).ToString().PadLeft(6)
                    : "  0.0"; // Default value if parsing fails
                gasData.CO = Math.Max(numericValue1, gasData.CO);
                string roundedValue2 = Math.Round(gasData.CO, 2).ToString().PadLeft(10);

                // Construct the updated value
                string updatedValue = currentValue.Substring(0, position) + roundedValue1;

                // Update the dictionary with the new value
                writeDictionary["CO"] = updatedValue;
            }
        }
        private void writeforCO2()
        {
            if (writeDictionary.ContainsKey("CO2"))
            {
                string currentValue = writeDictionary["CO2"];

                // Attempt to parse and round cellValues[5] and cellValues[6]
                string roundedValue1 = double.TryParse(cellValues[5], out double numericValue1)
                    ? Math.Round(numericValue1, 2).ToString().PadLeft(6)
                    : "  0.0"; // Default value if parsing fails
                gasData.CO2 = Math.Max(numericValue1, gasData.CO2);
                string roundedValue2 = Math.Round(gasData.CO2, 2).ToString().PadLeft(10);

                // Construct the updated value
                string updatedValue = currentValue.Substring(0, position) + roundedValue1;

                // Update the dictionary with the new value
                writeDictionary["CO2"] = updatedValue;
            }
        }
        private void writeforO2()
        {
            if (writeDictionary.ContainsKey("O2"))
            {
                string currentValue = writeDictionary["O2"];

                // Attempt to parse and round cellValues[7] and cellValues[8]
                string roundedValue1 = double.TryParse(cellValues[5], out double numericValue1)
                    ? Math.Round(numericValue1/100, 2).ToString().PadLeft(6)
                    : "  0.0"; // Default value if parsing fails
                gasData.O2 = Math.Max(numericValue1, gasData.O2);
                string roundedValue2 = Math.Round(gasData.O2, 2).ToString().PadLeft(10);

                // Construct the updated value
                string updatedValue = currentValue.Substring(0, position) + roundedValue1;

                // Update the dictionary with the new value
                writeDictionary["O2"] = updatedValue;
            }
        }
        private void writeforSPM()
        {
            if (writeDictionary.ContainsKey("SPM"))
            {
                string currentValue = writeDictionary["SPM"];

                // Attempt to parse and round cellValues[9] and cellValues[10]
                string roundedValue1 = double.TryParse(cellValues[6], out double numericValue1)
                    ? Math.Round(numericValue1, 2).ToString().PadLeft(6)
                    : "  0.0"; // Default value if parsing fails
                gasData.SPM = Math.Max(numericValue1, gasData.SPM);
                string roundedValue2 = Math.Round(gasData.SPM, 2).ToString().PadLeft(10);

                // Construct the updated value
                string updatedValue = currentValue.Substring(0, position) + roundedValue1;

                // Update the dictionary with the new value
                writeDictionary["SPM"] = updatedValue;
            }
        }
        private void writeforTemp()
        {
            if (writeDictionary.ContainsKey("Temp"))
            {
                string currentValue = writeDictionary["Temp"];

                // Attempt to parse and round cellValues[11] and cellValues[12]
                string roundedValue1 = double.TryParse(cellValues[7], out double numericValue1)
                    ? Math.Round(numericValue1, 2).ToString().PadLeft(6)
                    : "  0.0"; // Default value if parsing fails
                gasData.Temp = Math.Max(numericValue1, gasData.Temp);
                string roundedValue2 = Math.Round(gasData.Temp, 2).ToString().PadLeft(10);

                // Construct the updated value
                string updatedValue = currentValue.Substring(0, position) + roundedValue1;

                // Update the dictionary with the new value
                writeDictionary["Temp"] = updatedValue;
            }
        }
        private void writeforFlow()
        {
            if (writeDictionary.ContainsKey("Flow"))
            {
                string currentValue = writeDictionary["Flow"];

                // Attempt to parse and round cellValues[8] (Flow value)
                string roundedValue1 = double.TryParse(cellValues[8], out double numericValue1)
                    ? Math.Round(numericValue1, 2).ToString().PadLeft(6)
                    : "  0.0"; // Default value if parsing fails

                // Construct the updated value (append rounded value at the correct position)
                string updatedValue = currentValue.Substring(0, position) + roundedValue1;

                // Update the dictionary with the new value
                writeDictionary["Flow"] = updatedValue;
            }
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            ReadLastLineOfExcel();
            if (!string.IsNullOrEmpty(targetFilePath))
            {
                WriteDictionaryToFile(targetFilePath);
                ConvertCsvToExcel(targetFilePath); // Convert CSV to Excel after writing
            }
        }

        private void ConvertCsvToExcel(string csvFilePath)
        {
            // Generate Excel file path by replacing .csv with .xlsx
            string excelFilePath = Path.ChangeExtension(csvFilePath, ".xlsx");

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                workbook = excelApp.Workbooks.Add();
                worksheet = workbook.Sheets[1];

                string[] lines = File.ReadAllLines(csvFilePath);
                for (int row = 0; row < lines.Length; row++)
                {
                    string[] values = lines[row].Split(',');
                    for (int col = 0; col < values.Length; col++)
                    {
                        worksheet.Cells[row + 1, col + 1].Value = values[col];
                    }
                }

                workbook.SaveAs(excelFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook);
                MessageBox.Show($"CSV converted to Excel: {excelFilePath}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error converting CSV to Excel: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }

    }
}