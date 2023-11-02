using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace SejinTraceability
{
    public partial class ExportForm : Form
    {
        private string connectionString;
        private string selectedProject;
        private DateTime startDate;
        private ComboBox projectComboBox;
        private DateTimePicker startDatePicker;
        private DateTimePicker endDatePicker;
        private System.Windows.Forms.Button Exportbutton;
        private DateTime endDate;

        public ExportForm()
        {
            InitializeComponent();
        }

        private void projectComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedProject = projectComboBox.SelectedItem.ToString();
        }

        private void startDatePicker_ValueChanged(object sender, EventArgs e)
        {
            startDate = startDatePicker.Value;
        }

        private void endDatePicker_ValueChanged(object sender, EventArgs e)
        {
            endDate = endDatePicker.Value;
        }

        private void exportButton_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable dataTable = GetDataFromDatabase(selectedProject, startDate, endDate);

                if (dataTable.Rows.Count > 0)
                {
                    ExportToExcel(dataTable);
                    ShowSuccessMessage("Dane zostały pomyślnie wyeksportowane do Excela.");
                }
                else
                {
                    ShowErrorMessage("Brak danych do wyeksportowania.");
                }
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Błąd podczas eksportu danych: " + ex.Message);
            }
        }

        private System.Data.DataTable GetDataFromDatabase(string project, DateTime startDate, DateTime endDate)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT * FROM Archive WHERE project = @project AND date BETWEEN @startDate AND @endDate";
                using (SqlCommand cmd = new SqlCommand(query, connection))
                {
                    cmd.Parameters.AddWithValue("@project", project);
                    cmd.Parameters.AddWithValue("@startDate", startDate);
                    cmd.Parameters.AddWithValue("@endDate", endDate);

                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        adapter.Fill(dataTable);
                    }
                }
            }

            return dataTable;
        }

        private void ExportToExcel(System.Data.DataTable dataTable)
        {
            string excelFilePath = "exported_data.xlsx";
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Add();
            Worksheet worksheet = workbook.Worksheets[1] as Worksheet;

            for (int i = 1; i <= dataTable.Columns.Count; i++)
            {
                worksheet.Cells[1, i] = dataTable.Columns[i - 1].ColumnName;
            }

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataTable.Rows[i][j].ToString();
                }
            }

            workbook.SaveAs(excelFilePath);
            workbook.Close();
            Marshal.ReleaseComObject(workbook);
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
        }

        private void ShowSuccessMessage(string message)
        {
            MessageBox.Show(message, "Sukces", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowErrorMessage(string message)
        {
            MessageBox.Show(message, "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void InitializeComponent()
        {
            this.projectComboBox = new System.Windows.Forms.ComboBox();
            this.startDatePicker = new System.Windows.Forms.DateTimePicker();
            this.endDatePicker = new System.Windows.Forms.DateTimePicker();
            this.Exportbutton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // projectComboBox
            // 
            this.projectComboBox.FormattingEnabled = true;
            this.projectComboBox.Location = new System.Drawing.Point(12, 12);
            this.projectComboBox.Name = "projectComboBox";
            this.projectComboBox.Size = new System.Drawing.Size(200, 23);
            this.projectComboBox.TabIndex = 0;
            // 
            // startDatePicker
            // 
            this.startDatePicker.Location = new System.Drawing.Point(12, 62);
            this.startDatePicker.Name = "startDatePicker";
            this.startDatePicker.Size = new System.Drawing.Size(200, 23);
            this.startDatePicker.TabIndex = 1;
            // 
            // endDatePicker
            // 
            this.endDatePicker.Location = new System.Drawing.Point(12, 108);
            this.endDatePicker.Name = "endDatePicker";
            this.endDatePicker.Size = new System.Drawing.Size(200, 23);
            this.endDatePicker.TabIndex = 2;
            // 
            // Exportbutton
            // 
            this.Exportbutton.Location = new System.Drawing.Point(12, 149);
            this.Exportbutton.Name = "Exportbutton";
            this.Exportbutton.Size = new System.Drawing.Size(200, 23);
            this.Exportbutton.TabIndex = 3;
            this.Exportbutton.Text = "Export";
            this.Exportbutton.UseVisualStyleBackColor = true;
            this.Exportbutton.Click += new System.EventHandler(this.exportButton_Click);
            // 
            // ExportForm
            // 
            this.ClientSize = new System.Drawing.Size(219, 178);
            this.Controls.Add(this.Exportbutton);
            this.Controls.Add(this.endDatePicker);
            this.Controls.Add(this.startDatePicker);
            this.Controls.Add(this.projectComboBox);
            this.Name = "ExportForm";
            this.ResumeLayout(false);

        }

    }
}
