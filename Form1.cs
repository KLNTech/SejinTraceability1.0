using System;
using System.Data.SqlClient;
using System.Reactive.Linq;
using System.Reactive.Subjects;
using System.Reactive.Concurrency;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reactive;

namespace SejinTraceability
{
    public partial class TraceabilityForm : Form
    {
        private readonly string connectionString;
        private System.Windows.Forms.TextBox[] textBoxes;
        private int currentTextBoxIndex = 0;
        private readonly Subject<Unit> userInputSubject = new();
        private IDisposable inputSubscription;
        private const int MaxCharacterCount = 25;
        private const int MaxIdleTimeSeconds = 3;
        private bool projectSelectionPending = false;
        private bool isAutoMoveInProgress = false;
        private readonly object lockObject = new();
        private bool isFormOpened = false;
        private bool projectSelected = false;

        public TraceabilityForm()
        {
            InitializeComponent();
            connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["MyConnectionString"].ConnectionString;

            if (string.IsNullOrEmpty(connectionString))
            {
                throw new InvalidOperationException("ConnectionString property has not been initialized.");
            }

            InitializeFormTrace();
            textBoxes = new System.Windows.Forms.TextBox[] { textBoxtrace, textBoxtrace2, textBoxrackqty, textBoxrack, textBoxrack2 };
            var projectSelectionForm = new ProjectSelectionForm();
            projectSelectionForm.ProjectSelectedOnce += ProjectSelectionForm_ProjectSelectedOnce;
        }

        public void InitializeFormTrace()
        {
            foreach (var textBox in new[] { textBoxtrace, textBoxtrace2, textBoxPN, textBoxrackqty, textBoxrack, textBoxrack2 })
            {
                textBox.TextChanged += TextBox_TextChanged;
            }

            textBoxPN.Enabled = false;

            textBoxes = new System.Windows.Forms.TextBox[] { textBoxtrace, textBoxtrace2, textBoxrackqty, textBoxrack, textBoxrack2 };

            inputSubscription = userInputSubject
                .Throttle(TimeSpan.FromSeconds(MaxIdleTimeSeconds))
                .ObserveOn(SynchronizationContext.Current)
                .Subscribe(_ => ThrottleMoveToNextTextBox());
        }

        private void ShowSuccessMessage(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate { ShowSuccessMessage(message); }));
            }
            else
            {
                MessageBox.Show(message, "Sukces", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ShowErrorMessage(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate { ShowErrorMessage(message); }));
            }
            else
            {
                MessageBox.Show(message, "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string skipAutoMove = null;

        private void HandleTextChanged(string text)
        {
            Console.WriteLine($"HandleTextChanged: TextBoxIndex: {currentTextBoxIndex}, Text: {text}");

            if (text.Length == MaxCharacterCount)
            {
                if (textBoxes[currentTextBoxIndex] == textBoxtrace)
                {
                    string pn = text[13..];
                    textBoxPN.Text = pn;
                    userInputSubject.OnNext(Unit.Default);
                    textBoxPN.Enabled = false;
                    projectSelectionPending = false;

                    // Przesuñ do nastêpnego pola tylko, jeœli u¿ytkownik skoñczy³ wprowadzaæ ci¹g znaków
                    ThrottleMoveToNextTextBox();
                }
            }
            else if (text.Length != MaxCharacterCount && textBoxes[currentTextBoxIndex] == textBoxtrace)
            {
                projectSelectionPending = true;

                // Przesuñ do nastêpnego pola tylko, jeœli u¿ytkownik skoñczy³ wprowadzaæ ci¹g znaków
                ThrottleMoveToNextTextBox();
            }
        }

        private async Task MoveToNextTextBox()
        {
            Console.WriteLine($"MoveToNextTextBox: TextBoxIndex: {currentTextBoxIndex}");

            if (projectSelectionPending)
            {
                ShowProjectSelectionDialog();
                projectSelectionPending = false;
                return;
            }

            if (!string.IsNullOrEmpty(skipAutoMove))
            {
                skipAutoMove = null;
                return;
            }

            if (!projectSelected)
            {
                // Jeœli nie, to poczekaj
                await Task.Delay(500);
                return;
            }

            if (!isAutoMoveInProgress && string.IsNullOrWhiteSpace(textBoxes[currentTextBoxIndex].Text))
            {
                ShowMessageBox();
                return;
            }

            if (!isAutoMoveInProgress)
            {
                currentTextBoxIndex = (currentTextBoxIndex + 1) % textBoxes.Length;

                if (currentTextBoxIndex == 0 && textBoxes[currentTextBoxIndex] == textBoxrack2)
                {
                    Console.WriteLine("MoveToNextTextBox: Wprowadzono dane do ostatniego pola (textBoxrack2).");
                    ShowLastFieldMessageBox();
                    ThrottleMoveToNextTextBox();
                }
                else
                {
                    await Task.Delay(500);
                    SetActiveControl();
                }
            }
        }

        private async void ThrottleMoveToNextTextBox()
        {
            Console.WriteLine("ThrottleMoveToNextTextBox called");

            lock (lockObject)
            {
                if (isAutoMoveInProgress)
                {
                    return;
                }

                isAutoMoveInProgress = true;
            }

            Console.WriteLine($"Before MoveToNextTextBox: TextBoxIndex: {currentTextBoxIndex}");

            // Poczekaj na potwierdzenie, ¿e u¿ytkownik przesta³ wprowadzaæ dane
            await Task.Delay(TimeSpan.FromSeconds(MaxIdleTimeSeconds));

            // Dodatkowy warunek, aby unikn¹æ natychmiastowego pokazywania okna ProjectSelectionForm
            if (projectSelectionPending)
            {
                ShowProjectSelectionDialog();
                projectSelectionPending = false;
            }

            // SprawdŸ, czy dane zosta³y wprowadzone przez u¿ytkownika
            if (HasUserInput())
            {
                await MoveToNextTextBox();
            }

            Console.WriteLine($"After MoveToNextTextBox: TextBoxIndex: {currentTextBoxIndex}");

            lock (lockObject)
            {
                isAutoMoveInProgress = false;
            }
        }

        private bool HasUserInput()
        {
            // SprawdŸ, czy którykolwiek z TextBox ma wprowadzone dane
            return textBoxes.Any(tb => !string.IsNullOrWhiteSpace(tb.Text));
        }


        private void ShowMessageBox()
        {
            if (textBoxrackqty.InvokeRequired)
            {
                textBoxrackqty.Invoke(new MethodInvoker(delegate { ShowMessageBox(); }));
            }
            else
            {
                MessageBox.Show("WprowadŸ dane przed przejœciem dalej.");
            }
        }

        private void ShowLastFieldMessageBox()
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate { ShowLastFieldMessageBox(); }));
            }
            else
            {
                MessageBox.Show("Wprowadzono dane do ostatniego pola (textBoxrack2).");
            }
        }

        private void SetActiveControl()
        {
            if (textBoxes[currentTextBoxIndex].InvokeRequired)
            {
                textBoxes[currentTextBoxIndex].BeginInvoke((MethodInvoker)delegate
                {
                    ActiveControl = textBoxes[currentTextBoxIndex];
                });
            }
            else
            {
                ActiveControl = textBoxes[currentTextBoxIndex];
            }
        }

        private void TextBox_TextChanged(object sender, EventArgs e)
        {
            Console.WriteLine("Event called");
            System.Windows.Forms.TextBox textBox = (System.Windows.Forms.TextBox)sender;
            string text = textBox.Text;
            Console.WriteLine($"TextBox_TextChanged: Text: {text}");
            HandleTextChanged(text);
        }

        private void ProjectSelectionForm_ProjectSelectedOnce(object sender, string selectedProject)
        {
            projectSelected = true;
            isFormOpened = false;
        }

        private void ShowProjectSelectionDialog()
        {
            if (!projectSelectionPending || isFormOpened)
            {
                return;
            }

            isFormOpened = true;
            projectSelected = false; // Zresetuj flagê projectSelected
            ProjectSelectionForm projectSelectionForm = new ProjectSelectionForm();
            projectSelectionForm.ProjectSelectedOnce += ProjectSelectionForm_ProjectSelectedOnce;
            projectSelectionForm.ShowDialog();
        }

        private void PrintAndArchiveClick(object sender, EventArgs e)
        {
            OpenAndPrintExcelFileHandler(sender, e);
        }

        private void OpenAndPrintExcelFileHandler(object sender, EventArgs e)
        {
            string trace = textBoxtrace.Text;
            string trace2 = textBoxtrace2.Text;
            string rackQty = textBoxrackqty.Text;
            string rack = textBoxrack.Text;
            string rack2 = textBoxrack2.Text;

            if (trace.Length == 25)
            {
                string pn = trace[13..];
                string p_trace = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + rackQty + rack + pn;

                (string partname, string revValue, string barcodeValue) = GetPartNameRevAndBarcode(pn);
                string rev = revValue;
                string barcode = barcodeValue;

                try
                {
                    InsertRecord(pn, DateTime.Now.Date, DateTime.Now.TimeOfDay, rackQty, rack, rack2, trace, trace2, p_trace, barcode);
                    OpenAndPrintExcelFile(pn, DateTime.Now.Date, DateTime.Now.TimeOfDay, rackQty, rack, rack2, trace, trace2, p_trace, rev, barcode);

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


        private void InsertRecord(string pn, DateTime date, TimeSpan hour, string rackQty, string rack, string rack2, string trace, string trace2, string pTrace, string barcode)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                // Pobierz PartName i Rev z tabeli Database na podstawie PN
                string selectDatabaseQuery = "SELECT PartName, Rev FROM [Database] WHERE PN = @pn";

                using (SqlCommand selectDatabaseCmd = new SqlCommand(selectDatabaseQuery, connection))
                {
                    selectDatabaseCmd.Parameters.AddWithValue("@pn", pn);
                    connection.Open();
                    using (SqlDataReader reader = selectDatabaseCmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string partName = reader["PartName"].ToString();
                            string rev = reader["Rev"].ToString();

                            // Teraz, kiedy masz PartName i Rev, mo¿esz je wstawiæ do tabeli Archive
                            string insertQuery = "INSERT INTO Archive (PN, Date, Hour, RackQty, Rack, Rack2, Trace, Trace2, PTrace, Barcode, PartName, Rev) " +
                                                 "VALUES (@pn, @date, @hour, @rack_qty, @rack, @rack2, @trace, @trace2, @p_trace, @barcode, @part_name, @rev); " +
                                                 "SELECT CAST(SCOPE_IDENTITY() AS INT)";

                            using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection))
                            {
                                insertCmd.Parameters.AddWithValue("@pn", pn);
                                insertCmd.Parameters.AddWithValue("@date", date.ToString("MM/dd/yyyy"));
                                insertCmd.Parameters.AddWithValue("@hour", hour.ToString(@"hh\:mm\:ss"));
                                insertCmd.Parameters.AddWithValue("@rack_qty", rackQty);
                                insertCmd.Parameters.AddWithValue("@rack", rack);
                                insertCmd.Parameters.AddWithValue("@rack2", rack2);
                                insertCmd.Parameters.AddWithValue("@trace", trace);
                                insertCmd.Parameters.AddWithValue("@trace2", trace2);
                                insertCmd.Parameters.AddWithValue("@p_trace", pTrace);
                                insertCmd.Parameters.AddWithValue("@barcode", barcode);
                                insertCmd.Parameters.AddWithValue("@part_name", partName);
                                insertCmd.Parameters.AddWithValue("@rev", rev);

                                try
                                {
                                    // Wykonaj wstawienie do tabeli Archive
                                    int idTrace = (int)insertCmd.ExecuteScalar();
                                    ShowSuccessMessage($"Rekord zosta³ zarchiwizowany. id_trace = {idTrace}");
                                }
                                catch (Exception ex)
                                {
                                    ShowErrorMessage("B³¹d podczas archiwizacji: " + ex.Message);
                                }
                            }
                        }
                        else
                        {
                            // Obs³u¿ przypadek, gdy nie znaleziono informacji dla danego PN w tabeli Database
                            ShowErrorMessage($"Brak informacji w tabeli Database dla PN: {pn}");
                        }
                    }
                }
            }
        }




        private (string partName, string revValue, string barcodeValue) GetPartNameRevAndBarcode(string pn)
        {
            string partName = string.Empty;
            string rev = string.Empty;
            string barcode = string.Empty;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Pobierz PartName, Rev z tabeli Database na podstawie PN
                string selectDatabaseQuery = "SELECT TOP 1 [PartName], [Rev] FROM [Database] WHERE PN = @pn";

                using (SqlCommand cmd = new SqlCommand(selectDatabaseQuery, connection))
                {
                    cmd.Parameters.AddWithValue("@pn", pn);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            partName = reader["PartName"].ToString();
                            rev = reader["Rev"].ToString();
                        }
                    }
                }

                // Pobierz ostatni¹ wartoœæ kolumny "Barcode" z tabeli "Archive" na podstawie PN
                string selectArchiveBarcodeQuery = "SELECT TOP 1 [Barcode] FROM [Archive] WHERE PN = @pn ORDER BY Date DESC, Hour DESC";

                using (SqlCommand cmd = new SqlCommand(selectArchiveBarcodeQuery, connection))
                {
                    cmd.Parameters.AddWithValue("@pn", pn);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            barcode = reader["Barcode"].ToString();
                        }
                    }
                }
            }

            // Generuj now¹ wartoœæ kolumny "Barcode" na podstawie poprzedniej
            string firstPart = barcode[..7];
            string secondPart = barcode[7..];

            if (int.TryParse(secondPart, out int secondPartNumber))
            {
                secondPartNumber++;
                string incrementedSecondPart = secondPartNumber.ToString("D6");
                barcode = firstPart + incrementedSecondPart;
            }

            return (partName, rev, barcode);
        }



        private static void OpenAndPrintExcelFile(string pn, DateTime date, TimeSpan hour, string rackQty, string rack, string rack2, string trace, string trace2, string p_trace, string rev, string barcode)
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
                    worksheet.Range["I2"].Value = "VW"; // Dla 25 znaków w textBoxtrace
                    worksheet.Range["G10"].Value = rev;
                    worksheet.Range["I6"].Value = barcode;
                    worksheet.Range["G26"].Value = p_trace;
                    worksheet.Range["A18"].Value = rackQty;
                    worksheet.Range["E18"].Value = date.ToString("yyyy-MM-dd");
                    worksheet.Range["C21"].Value = hour.ToString(@"hh\:mm\:ss");
                    worksheet.Range["A14"].Value = p_trace;
                    worksheet.Range["A26"].Value = trace2; // Nowe pole trace2

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
