using System;
using System.Data.SqlClient;
using System.Reactive.Linq;
using System.Reactive.Subjects;
using System.Reactive.Concurrency;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reactive;

namespace SejinTraceability
{
    public partial class straceabilitysystem : Form
    {
        private string connectionString;
        private System.Windows.Forms.TextBox[] textBoxes;
        private int currentTextBoxIndex = 0;
        private Subject<Unit> userInputSubject = new Subject<Unit>();
        private IDisposable inputSubscription;
        private string[] projects = { "PM1", "PM2", "PM3" };

        public straceabilitysystem()
        {
            InitializeComponent();
            connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["MyConnectionString"].ConnectionString;

            if (string.IsNullOrEmpty(connectionString))
            {
                throw new InvalidOperationException("ConnectionString property has not been initialized.");
            }

            InitializeFormTrace();
        }

        public void InitializeFormTrace()
        {
            textBoxtrace.TextChanged += TextBoxTrace_TextChanged;
            textBoxtrace2.TextChanged += TextBoxTrace2_TextChanged;
            textBoxPN.TextChanged += TextBoxPN_TextChanged;
            textBoxrackqty.TextChanged += TextBox_TextChanged;
            textBoxrack.TextChanged += TextBox_TextChanged;
            textBoxrack2.TextChanged += TextBox_TextChanged;

            textBoxes = new System.Windows.Forms.TextBox[] { textBoxtrace, textBoxtrace2, textBoxPN, textBoxrackqty, textBoxrack, textBoxrack2 };

            inputSubscription = userInputSubject
                .Throttle(TimeSpan.FromSeconds(1))
                .ObserveOn(SynchronizationContext.Current)
                .Subscribe(_ => MoveToNextTextBox());
        }
        private void ShowSuccessMessage(string message)
        {
            MessageBox.Show(message, "Sukces", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowErrorMessage(string message)
        {
            MessageBox.Show(message, "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }


        private void HandleTextChanged(string text)
        {
            if (text.Length == 25 && textBoxes[currentTextBoxIndex] == textBoxtrace)
            {
                string pn = text.Substring(13);
                textBoxPN.Text = pn;
                userInputSubject.OnNext(Unit.Default);
                MoveToNextTextBox();
            }
            else if (text.Length != 25 && textBoxes[currentTextBoxIndex] == textBoxtrace)
            {
                ShowProjectSelectionDialog();
            }
            else
            {
                MoveToNextTextBox();
            }
        }

            private void MoveToNextTextBox()
        {
            if (currentTextBoxIndex < textBoxes.Length - 1)
            {
                currentTextBoxIndex++;
            }
            else
            {
                currentTextBoxIndex = 0;
                MessageBox.Show("Wprowadzono dane do ostatniego pola (textBoxrack2).");
            }

            ActiveControl = textBoxes[currentTextBoxIndex];
        }

        private void TextBoxTrace_TextChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox textBox = (System.Windows.Forms.TextBox)sender;
            string text = textBox.Text;
            HandleTextChanged(text);
        }

        private void TextBoxTrace2_TextChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox textBox = (System.Windows.Forms.TextBox)sender;
            string text = textBox.Text;
            HandleTextChanged(text);
        }

        private void TextBoxPN_TextChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox textBox = (System.Windows.Forms.TextBox)sender;
            string text = textBox.Text;
            HandleTextChanged(text);
        }

        private void TextBox_TextChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox textBox = (System.Windows.Forms.TextBox)sender;
            string text = textBox.Text;
            HandleTextChanged(text);
        }

        private void ShowProjectSelectionDialog()
        {
            var projectSelectionForm = new ProjectSelectionForm();
            projectSelectionForm.ProjectSelected += (sender, selectedProject) =>
            {
                textBoxPN.Text = selectedProject;
                userInputSubject.OnNext(Unit.Default);
            };
            projectSelectionForm.ShowDialog();
        }

        private void PrintAndArchiveClick(object sender, EventArgs e)
        {
            OpenAndPrintExcelFileHandler(sender, e);
        }

        private void OpenAndPrintExcelFileHandler(object sender, EventArgs e)
        {
            string trace = textBoxtrace.Text;
            string rackQty = textBoxrackqty.Text;
            string rack = textBoxrack.Text;

            if (trace.Length == 25)
            {
                string pn = trace.Substring(13);
                string p_trace = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + rackQty + rack + pn;

                (string revValue, string barcodeValue) result = GetRevAndBarcode(pn);
                string rev = result.revValue;
                string barcode = result.barcodeValue;

                try
                {
                    InsertRecord(pn, DateTime.Now.Date, DateTime.Now.TimeOfDay, rackQty, rack, trace, p_trace, rev, barcode);
                    OpenAndPrintExcelFile(pn, DateTime.Now.Date, DateTime.Now.TimeOfDay, rackQty, rack, trace, p_trace, rev, barcode);

                    ShowSuccessMessage("Plik Excel zosta³ otwarty i wydrukowany.");
                }
                catch (Exception ex)
                {
                    ShowErrorMessage("B³¹d: " + ex.Message);
                }
                finally
                {
                    // Dzia³ania do wykonania po zakoñczeniu próby otwarcia i wydrukowania pliku Excel.
                }
            }
            else if (trace.Length > 25 || !string.IsNullOrEmpty(rack))
            {
                // Obs³uga przypadku dla instrukcji "traceability" dla VW
                // Tutaj mo¿esz dodaæ kod obs³uguj¹cy instrukcjê "traceability"
                // oraz archiwizacjê wyników tej instrukcji
            }
            else
            {
                ShowErrorMessage("B³¹d: Nieprawid³owa d³ugoœæ ci¹gu lub brak danych.");
            }
        }

        private void InsertRecord(string pn, DateTime date, TimeSpan hour, string rackQty, string rack, string trace, string pTrace, string rev, string barcode)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string insertQuery = "INSERT INTO Archive (pn, date, hour, rack_qty, rack, trace, p_trace, rev, barcode) " +
                                     "VALUES (@pn, @date, @hour, @rack_qty, @rack, @trace, @p_trace, @rev, @barcode)";

                using (SqlCommand cmd = new SqlCommand(insertQuery, connection))
                {
                    cmd.Parameters.AddWithValue("@pn", pn);
                    cmd.Parameters.AddWithValue("@date", date);
                    cmd.Parameters.AddWithValue("@hour", hour);
                    cmd.Parameters.AddWithValue("@rack_qty", rackQty);
                    cmd.Parameters.AddWithValue("@rack", rack);
                    cmd.Parameters.AddWithValue("@trace", trace);
                    cmd.Parameters.AddWithValue("@p_trace", pTrace);
                    cmd.Parameters.AddWithValue("@rev", rev);
                    cmd.Parameters.AddWithValue("@barcode", barcode);

                    try
                    {
                        connection.Open();
                        cmd.ExecuteNonQuery();
                        ShowSuccessMessage("Rekord zosta³ zarchiwizowany.");
                    }
                    catch (Exception ex)
                    {
                        ShowErrorMessage("B³¹d podczas archiwizacji: " + ex.Message);
                    }
                }
            }
        }

        private (string revValue, string barcodeValue) GetRevAndBarcode(string pn)
        {
            string rev = string.Empty;
            string barcode = string.Empty;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT TOP 1 barcode FROM Archive ORDER BY date DESC, hour DESC";
                using (SqlCommand cmd = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            barcode = reader["barcode"].ToString();
                        }
                    }
                }
            }

            string firstPart = barcode.Substring(0, 7);
            string secondPart = barcode.Substring(7);

            if (int.TryParse(secondPart, out int secondPartNumber))
            {
                secondPartNumber++;
                string incrementedSecondPart = secondPartNumber.ToString("D6");
                barcode = firstPart + incrementedSecondPart;
            }

            return (rev, barcode);
        }

        private void OpenAndPrintExcelFile(string pn, DateTime date, TimeSpan hour, string rackQty, string rack, string trace, string p_trace, string rev, string barcode)
        {
            string currentDirectory = Directory.GetCurrentDirectory();
            string excelFileName = "label.xlsx";
            string excelFilePath = Path.Combine(currentDirectory, excelFileName);

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            try
            {
                Workbook workbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly: false, UpdateLinks: false);
                Worksheet worksheet = null;

                foreach (Worksheet sheet in workbook.Sheets)
                {
                    if (sheet.Name == "label")
                    {
                        worksheet = sheet;
                        break;
                    }
                }

                if (worksheet != null)
                {
                    worksheet.Range["A7"].Value = pn;
                    worksheet.Range["I2"].Value = "PM";
                    worksheet.Range["G10"].Value = rev;
                    worksheet.Range["I6"].Value = barcode;
                    worksheet.Range["A14"].Value = p_trace;
                    worksheet.Range["A18"].Value = rackQty;
                    worksheet.Range["E18"].Value = date.ToString("yyyy-MM-dd");
                    worksheet.Range["C21"].Value = hour.ToString(@"hh\:mm\:ss");

                    worksheet.PrintOut();
                    workbook.Close(false, excelFilePath, Type.Missing);
                }
                else
                {
                    Console.WriteLine("Nie znaleziono arkusza o nazwie 'label'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("B³¹d: " + ex.Message);
            }
            finally
            {
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }

        private void ExportButtonClick(object sender, EventArgs e)
        {
            var exportForm = new ExportForm();
            exportForm.ShowDialog();
        }
    }
}
