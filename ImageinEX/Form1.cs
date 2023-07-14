using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ImageinEX
{
    public partial class Form1 : Form
    {
        private string selectedFilePath = string.Empty;
        private string selectedSheetName = string.Empty;
        private string selectedColumnName = string.Empty;
        private string selectedFolderPath = string.Empty;
        private DataTable dataTable;

        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Initialize the dropdown list of sheet names
            sheetComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            sheetComboBox.Enabled = false;

            // Initialize the DataGridView
            dataTable = new DataTable();
            dataGridView.DataSource = dataTable;
        }

        private void browseFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
            openFileDialog.Title = "اختر ملف الإكسل";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = openFileDialog.FileName;
                filePathTextBox.Text = selectedFilePath;
                LoadSheetNames(selectedFilePath);
            }
        }

        private void LoadSheetNames(string filePath)
        {
            sheetComboBox.Items.Clear();
            columnComboBox.Items.Clear();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheets = package.Workbook.Worksheets;
                foreach (var worksheet in worksheets)
                {
                    sheetComboBox.Items.Add(worksheet.Name);
                }
            }

            if (sheetComboBox.Items.Count > 0)
            {
                sheetComboBox.Enabled = true;
                sheetComboBox.SelectedIndex = 0;
                columnComboBox.Enabled = true;
                searchButton.Enabled = true;
            }
            else
            {
                sheetComboBox.Enabled = false;
                columnComboBox.Enabled = false;
                searchButton.Enabled = false;
            }
        }

        private void sheetComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedSheetName = sheetComboBox.SelectedItem.ToString();
            UpdateDataTable();
            LoadColumnNames(selectedSheetName);
            UpdateDataGridView();
        }

        private void UpdateDataTable()
        {
            dataTable.Clear();

            using (ExcelPackage package = new ExcelPackage(new FileInfo(selectedFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[selectedSheetName];

                int rows = worksheet.Dimension.Rows;
                int columns = worksheet.Dimension.Columns;

                for (int i = 1; i <= columns; i++)
                {
                    string columnName = worksheet.Cells[1, i].Value?.ToString();
                    if (!dataTable.Columns.Contains(columnName))
                    {
                        dataTable.Columns.Add(columnName);
                    }
                }

                for (int row = 2; row <= rows; row++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= columns; col++)
                    {
                        string columnName = worksheet.Cells[1, col].Value?.ToString();
                        string cellValue = worksheet.Cells[row, col].Value?.ToString();

                        if (dataTable.Columns.Contains(columnName))
                        {
                            dataRow[columnName] = cellValue;
                        }
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }
        }

        private void LoadColumnNames(string sheetName)
        {
            columnComboBox.Items.Clear();

            using (ExcelPackage package = new ExcelPackage(new FileInfo(selectedFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

                int columns = worksheet.Dimension.Columns;

                for (int col = 1; col <= columns; col++)
                {
                    string columnName = worksheet.Cells[1, col].Value?.ToString();
                    columnComboBox.Items.Add(columnName);
                }
            }
        }

        private void columnComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedColumnName = columnComboBox.SelectedItem?.ToString();
            UpdateDataGridView();
        }

        private void UpdateDataGridView()
        {
            if (!string.IsNullOrEmpty(selectedColumnName))
            {
                DataView dataView = dataTable.DefaultView;
                dataView.RowFilter = $"[{selectedColumnName}] <> ''";
                dataGridView.DataSource = dataView;
            }
            else
            {
                dataGridView.DataSource = dataTable;
            }
        }

        private void browseFolderButton_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFolderPath = folderBrowserDialog.SelectedPath;
                    folderPathTextBox.Text = selectedFolderPath;
                }
            }
        }

        private void searchButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(selectedColumnName) || string.IsNullOrEmpty(selectedFolderPath))
            {
                MessageBox.Show("Please select a column and a folder path.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DataTable newDataTable = dataTable.Clone();
            newDataTable.Columns.Add("IMAGE", typeof(string));

            foreach (DataRow row in dataTable.Rows)
            {
                string cellValue = row[selectedColumnName].ToString();
                List<string> matchingFolders = Directory.GetDirectories(selectedFolderPath, $"{cellValue}*")
                                                        .Select(folderPath => new DirectoryInfo(folderPath).Name)
                                                        .ToList();

                if (matchingFolders.Count > 0)
                {
                    foreach (string folderName in matchingFolders)
                    {
                        DataRow newRow = newDataTable.NewRow();
                        newRow.ItemArray = row.ItemArray;
                        newRow["IMAGE"] = folderName;
                        newDataTable.Rows.Add(newRow);
                    }
                }
                else
                {
                    DataRow newRow = newDataTable.NewRow();
                    newRow.ItemArray = row.ItemArray;
                    newRow["IMAGE"] = "";
                    newDataTable.Rows.Add(newRow);
                }
            }

            dataTable = newDataTable;
            UpdateDataGridView();
            saveButton.Enabled = true;
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            if (dataTable.Columns.Contains("IMAGE"))
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo fileInfo = new FileInfo(saveFileDialog.FileName);
                        using (ExcelPackage package = new ExcelPackage(fileInfo))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(selectedSheetName);

                            int rows = dataTable.Rows.Count;
                            int columns = dataTable.Columns.Count;

                            for (int col = 1; col <= columns; col++)
                            {
                                worksheet.Cells[1, col].Value = dataTable.Columns[col - 1].ColumnName;
                            }

                            for (int row = 2; row <= rows + 1; row++)
                            {
                                for (int col = 1; col <= columns; col++)
                                {
                                    worksheet.Cells[row, col].Value = dataTable.Rows[row - 2][col - 1].ToString();
                                }
                            }

                            package.Save();
                            MessageBox.Show("Table saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("There is no IMAGE column to save.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
