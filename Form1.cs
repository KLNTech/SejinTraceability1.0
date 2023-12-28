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
using ZXing;
using ZXing.Common;
using ZXing.Rendering;
using System.Net.NetworkInformation;
using stdole;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using ZXing.QrCode;
using DocumentFormat.OpenXml.Drawing;
using System.Diagnostics;
using System.Reflection;



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
        private const int MaxIdleTimeSeconds = 4;
        private bool projectSelectionPending = false;
        private bool isAutoMoveInProgress = false;
        private readonly object lockObject = new();
        private bool isFormOpened = false;
        private bool projectSelected = false;
        private string rev;
        private string rackQty;
        private string rack2;
        private string pn;
        private bool isHandlingTextChanged = false;

        public TraceabilityForm()
        {
            InitializeComponent();
            connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["MyConnectionString"].ConnectionString;
            SynchronizationContext syncContext = SynchronizationContext.Current;

            if (syncContext == null)
            {
                throw new InvalidOperationException("SynchronizationContext is not available.");
            }

            if (string.IsNullOrEmpty(connectionString))
            {
                throw new InvalidOperationException("ConnectionString property has not been initialized.");
            }

            CheckDatabaseConnection();  // SprawdŸ po³¹czenie z baz¹ danych
            CheckExcelFileAvailability(); // SprawdŸ dostêpnoœæ pliku Excel
            InitializeFormTrace();
            textBoxes = new System.Windows.Forms.TextBox[] { textBoxtrace, textBoxtrace2, textBoxrackqty, textBoxrack, textBoxrack2 };
            var projectSelectionForm = new ProjectSelectionForm();
            projectSelectionForm.ProjectSelectedOnce += ProjectSelectionForm_ProjectSelectedOnce;

            // Dodaj subskrypcjê po utworzeniu obiektu ProjectSelectionForm
            inputSubscription = userInputSubject
                .Throttle(TimeSpan.FromSeconds(MaxIdleTimeSeconds))
                .ObserveOn(syncContext)  // Obserwuj na g³ównym w¹tku
                .Subscribe(_ => ThrottleMoveToNextTextBox());
        }

        public void InitializeFormTrace()
        {
            foreach (var textBox in new[] { textBoxtrace, textBoxtrace2, textBoxPN, textBoxrackqty, textBoxrack, textBoxrack2 })
            {
                textBox.TextChanged += TextBox_TextChanged;
            }

            

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

        private void CheckDatabaseConnection()
        {
            using (SqlConnection connection = new(connectionString))
            {
                try
                {
                    connection.Open();
                    ShowSuccessMessage("Po³¹czenie z baz¹ danych zosta³o nawi¹zane.");
                }
                catch (SqlException ex)
                {
                    ShowErrorMessage("B³¹d po³¹czenia z baz¹ danych: " + ex.Message);
                }
            }
        }
        private void CheckExcelFileAvailability()
        {
            // Pobierz pe³n¹ œcie¿kê do pliku wykonywalnego aplikacji
            string executablePath = Assembly.GetExecutingAssembly().Location;
            string executableDirectory = System.IO.Path.GetDirectoryName(executablePath);

            string labelDirectory = System.IO.Path.Combine(executableDirectory, "Label"); // Folder "Label" w tym samym katalogu, co plik wykonywalny
            string excelFileName = "label.xlsx";
            string excelFilePath = System.IO.Path.Combine(labelDirectory, excelFileName);

            if (File.Exists(excelFilePath))
            {
                ShowSuccessMessage("Plik Excel jest dostêpny.");
            }
            else
            {
                ShowErrorMessage("Plik Excel nie jest dostêpny w lokalizacji: " + excelFilePath);
            }
        }


        private string skipAutoMove = null;

        private void HandleTextChanged(string text)
        {
            string message = $"HandleTextChanged: TextBoxIndex: {currentTextBoxIndex}, Text: {text}";

            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate
                {
                    MessageBox.Show(message);
                }));
            }
            else
            {
                MessageBox.Show(message);
            }

            if (text.Length == MaxCharacterCount)
            {
                if (textBoxes[currentTextBoxIndex] == textBoxtrace)
                {
                    string pn = text[13..];
                    textBoxPN.Text = pn;
                    userInputSubject.OnNext(Unit.Default);
                    textBoxPN.Enabled = false;

                    // Przesuñ do nastêpnego pola od razu, gdy u¿ytkownik skoñczy wprowadzaæ ci¹g znaków
                    MoveToNextTextBox();
                }
            }
            else if (text.Length != MaxCharacterCount && textBoxes[currentTextBoxIndex] == textBoxtrace)
            {
                // Przesuñ do nastêpnego pola tylko, jeœli u¿ytkownik skoñczy³ wprowadzaæ ci¹g znaków
                MoveToNextTextBox();
            }
        }


        private async Task MoveToNextTextBox()
        {
            try
            {
                Debug.WriteLine($"MoveToNextTextBox: TextBoxIndex: {currentTextBoxIndex}");

                if (projectSelectionPending)
                {
                    ShowProjectSelectionDialog();
                    projectSelectionPending = false;
                    return;
                }

                if (!string.IsNullOrEmpty(skipAutoMove))
                {
                    Debug.WriteLine("MoveToNextTextBox: No user input, showing message box.");
                    skipAutoMove = null;
                    return;
                }

                if (!projectSelected)
                {
                    // Poczekaj na potwierdzenie wyboru projektu
                    await Task.Delay(200);
                    return;
                }

                if (!isAutoMoveInProgress && string.IsNullOrWhiteSpace(textBoxes[currentTextBoxIndex].Text))
                {
                    Debug.WriteLine("MoveToNextTextBox: No user input, showing message box.");
                    ShowMessageBox();
                    return;
                }

                if (!isAutoMoveInProgress)
                {
                    Debug.WriteLine("MoveToNextTextBox: Changing currentTextBoxIndex and setting active control.");

                    if (InvokeRequired)
                    {
                        Invoke(new MethodInvoker(() => currentTextBoxIndex = (currentTextBoxIndex + 1) % textBoxes.Length));
                    }
                    else
                    {
                        currentTextBoxIndex = (currentTextBoxIndex + 1) % textBoxes.Length;
                    }

                    if (currentTextBoxIndex == 0 && textBoxes[currentTextBoxIndex] == textBoxrack2)
                    {
                        Debug.WriteLine("MoveToNextTextBox: Wprowadzono dane do ostatniego pola (textBoxrack2).");
                        ShowLastFieldMessageBox();
                    }

                    await Task.Delay(500);

                    if (InvokeRequired)
                    {
                        Invoke(new MethodInvoker(() => SetActiveControl()));
                    }
                    else
                    {
                        SetActiveControl();
                    }

                    Debug.WriteLine($"MoveToNextTextBox: ActiveControl set to TextBoxIndex: {currentTextBoxIndex}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"MoveToNextTextBox Error: {ex.Message}");
            }
        }




        private async void ThrottleMoveToNextTextBox()
        {
            await ExecuteThrottleMove();
            Debug.WriteLine("ThrottleMoveToNextTextBox called");

            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(() => ExecuteThrottleMove()));
            }
            else
            {
                ExecuteThrottleMove();
            }
        }

        private async void ExecuteThrottleMove()
        {
            if (Monitor.TryEnter(lockObject))
            {
                try
                {
                    if (isAutoMoveInProgress)
                    {
                        Debug.WriteLine("ThrottleMoveToNextTextBox: Auto move already in progress, returning.");
                        return;
                    }

                    isAutoMoveInProgress = true;
                }
                finally
                {
                    Monitor.Exit(lockObject);
                }

                Debug.WriteLine($"Before MoveToNextTextBox: TextBoxIndex: {currentTextBoxIndex}");

                try
                {
                    // Poczekaj na potwierdzenie, ¿e u¿ytkownik przesta³ wprowadzaæ dane
                    await Task.Delay(TimeSpan.FromSeconds(MaxIdleTimeSeconds));

                    // SprawdŸ, czy dane zosta³y wprowadzone przez u¿ytkownika
                    if (HasUserInput())
                    {
                        await MoveToNextTextBox();
                    }

                    Debug.WriteLine($"After MoveToNextTextBox: TextBoxIndex: {currentTextBoxIndex}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"ThrottleMoveToNextTextBox Error: {ex.Message}");
                }
                finally
                {
                    Monitor.Enter(lockObject);
                    isAutoMoveInProgress = false;
                    Monitor.Exit(lockObject);
                }
            }
            else
            {
                Debug.WriteLine("ThrottleMoveToNextTextBox: Could not enter critical section, returning.");
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
            if (textBoxes[currentTextBoxIndex] != null)
            {
                if (textBoxes[currentTextBoxIndex].InvokeRequired)
                {
                    textBoxes[currentTextBoxIndex].Invoke(new MethodInvoker(() =>
                    {
                        ActiveControl = textBoxes[currentTextBoxIndex];
                        Debug.WriteLine($"SetActiveControl: ActiveControl set to TextBoxIndex: {currentTextBoxIndex}");
                    }));
                }
                else
                {
                    ActiveControl = textBoxes[currentTextBoxIndex];
                    Debug.WriteLine($"SetActiveControl: ActiveControl set to TextBoxIndex: {currentTextBoxIndex}");
                }
            }
        }

        private void TextBox_TextChanged(object sender, EventArgs e)
        {
            if (isHandlingTextChanged)
            {
                return;
            }
            isHandlingTextChanged = true;
            Debug.WriteLine("Event called");
            System.Windows.Forms.TextBox textBox = (System.Windows.Forms.TextBox)sender;
            string text = textBox.Text;
            Debug.WriteLine($"TextBox_TextChanged: Text: {text}");
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
            ProjectSelectionForm projectSelectionForm = new();
            projectSelectionForm.ProjectSelectedOnce += ProjectSelectionForm_ProjectSelectedOnce;
            projectSelectionForm.ShowDialog();
        }


        private string GenerateLotCode(DateTime date)
        {
            // Logika generowania kodu lotu na podstawie daty
            char yearCode = (char)('A' + (date.Year - 2023) % 26);
            char monthCode = (char)('A' + date.Month - 1);

            // Jeœli przekroczono 25 liter alfabetu, zaczynamy u¿ywaæ cyfr (1 dla literki A, 2 dla B, itd.)
            char dayCode = date.Day <= 25 ? (char)('A' + date.Day - 1) : (char)('1' + date.Day - 26);

            return $"{yearCode}{monthCode}{dayCode}";
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
                rev = revValue;
                string barcode = barcodeValue;

                try
                {
                    GenerateAndSaveQRCode(trace, trace2, pn, rev, rackQty, barcode);
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

        private void GenerateBarcodeAndSave(string data, string barcodeValue, string fileName)
        {
            // Utwórz obiekt BarcodeWriter dla kodu kreskowego
            BarcodeWriter<Bitmap> barcodeWriter = new();
            barcodeWriter.Format = BarcodeFormat.CODE_128;

            // Do³¹cz wartoœæ zmiennej barcode do danych kodu kreskowego
            string fullData = $"{data} {barcodeValue}";

            // Utwórz obraz kodu kreskowego
            Bitmap barcodeBitmap = barcodeWriter.Write(fullData);

            // Utwórz katalog "Barcode", jeœli nie istnieje
            string barcodeDirectory = System.IO.Path.GetDirectoryName(fileName);
            if (!Directory.Exists(barcodeDirectory))
            {
                Directory.CreateDirectory(barcodeDirectory);
            }

            // Zapisz obraz kodu kreskowego do pliku
            barcodeBitmap.Save(fileName, ImageFormat.Png);
        }

        private void GenerateQRCodeAndSave(string qrText, string fileName)
        {
            // Utwórz obiekt BarcodeWriter z odpowiednimi parametrami typu
            BarcodeWriter<Bitmap> barcodeWriter = new();
            barcodeWriter.Format = BarcodeFormat.QR_CODE;
            barcodeWriter.Options = new ZXing.Common.EncodingOptions
            {
                Width = 300,
                Height = 300,
                Margin = 0
            };

            // Utwórz obraz kodu QR
            Bitmap qrBitmap = barcodeWriter.Write(qrText);

            // Zapisz obraz kodu QR do pliku
            qrBitmap.Save(fileName, ImageFormat.Png);
        }

        private void GenerateAndSaveQRCode(string trace, string trace2, string pn, string rev, string rackQty, string barcode)
        {
            if (!string.IsNullOrEmpty(trace) || !string.IsNullOrEmpty(trace2))
            {
                string qrText = string.Empty;

                if (!string.IsNullOrEmpty(trace) && !string.IsNullOrEmpty(trace2))
                {
                    qrText = $"[)>06:AS\"barcode\":PN\"{pn}\":QT\"{rackQty}.000\":RV\"{rev}\":DM\"{DateTime.Now.ToString("ddMMyy")}\":SPHS:PO:LT\"{GenerateLotCode(DateTime.Now)}\":WT\"{trace}\" / \"{trace2}\":PT\"{DateTime.Now.ToString("dd.MM.yy")} {DateTime.Now.TimeOfDay}\"/#{rack} / #{rack2}/{pn}:*[]\"";
                }
                else if (!string.IsNullOrEmpty(trace))
                {
                    qrText = $"[)>06:AS\"barcode\":PN\"{pn}\":QT\"{rackQty}.000\":RV\"{rev}\":DM\"{DateTime.Now.ToString("ddMMyy")}\":SPHS:PO:LT\"{GenerateLotCode(DateTime.Now)}\":WT\"{trace}\":PT\"{DateTime.Now.ToString("dd.MM.yy")} {DateTime.Now.TimeOfDay}\"/#{rack} / #{rack2}/{pn}:*[]\"";
                }

                // U¿yj wczeœniej zdefiniowanych metod do generowania i zapisywania kodu kreskowego oraz QR
                GenerateBarcodeAndSave(trace, barcode, "Barcode\\" + "Barcode.png");
                GenerateQRCodeAndSave(qrText, "QRCode\\" + "QRCode.png");

            }
        }


        private void InsertRecord(string pn, DateTime date, TimeSpan hour, string rackQty, string rack, string rack2, string trace, string trace2, string pTrace, string barcode)
        {
            using (SqlConnection connection = new(connectionString))
            {
                // Pobierz PartName i Rev z tabeli Database na podstawie PN
                string selectDatabaseQuery = "SELECT PartName, Rev FROM [Database] WHERE PN = @pn";

                using (SqlCommand selectDatabaseCmd = new(selectDatabaseQuery, connection))
                {
                    selectDatabaseCmd.Parameters.AddWithValue("@pn", pn);
                    connection.Open();

                    using (SqlDataReader reader = selectDatabaseCmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string partName = reader["PartName"].ToString();
                            string rev = reader["Rev"].ToString();

                            // Zamknij DataReader, poniewa¿ ju¿ uzyskaliœmy potrzebne dane
                            reader.Close();

                            // Teraz, kiedy masz PartName i Rev, mo¿esz je wstawiæ do tabeli Archive
                            string insertQuery = "INSERT INTO Archive (PN, Date, Hour, RackQty, Rack, Rack2, Trace, Trace2, PTrace, Barcode, PartName, Rev) " +
                                                 "VALUES (@pn, @date, @hour, @rack_qty, @rack, @rack2, @trace, @trace2, @p_trace, @barcode, @part_name, @rev); " +
                                                 "SELECT CAST(SCOPE_IDENTITY() AS INT)";

                            using (SqlCommand insertCmd = new(insertQuery, connection))
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

            using (SqlConnection connection = new(connectionString))
            {
                connection.Open();

                // Pobierz PartName, Rev z tabeli Database na podstawie PN
                string selectDatabaseQuery = "SELECT TOP 1 [PartName], [Rev] FROM [Database] WHERE PN = @pn";

                using (SqlCommand cmd = new(selectDatabaseQuery, connection))
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

                using (SqlCommand cmd = new(selectArchiveBarcodeQuery, connection))
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
            string labelDirectory = "Label"; // Nowa lokalizacja dla pliku Excela

            string excelFilePath = System.IO.Path.Combine(currentDirectory, labelDirectory, excelFileName);


            Microsoft.Office.Interop.Excel.Application excelApp = new();

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
                    Debug.WriteLine("Nie znaleziono arkusza o nazwie 'label'.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("B³¹d: " + ex.Message);
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
