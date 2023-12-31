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
using ZXing.Rendering;



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
        private readonly object lockObject = new();             
        private string rev;
        private string rackQty;
        private string rack2;
        private string pn;        
        private const int ThrottleTimeSeconds = 1;


        public TraceabilityForm()
        {
            InitializeComponent();
            connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["MyConnectionString"].ConnectionString;
            CheckDatabaseConnection();  // Sprawd� po��czenie z baz� danych
            CheckExcelFileAvailability(); // Sprawd� dost�pno�� pliku Excel
            textBoxes = new System.Windows.Forms.TextBox[] { textBoxtrace, textBoxtrace2, textBoxPN, textBoxrackqty, textBoxrack, textBoxrack2 };
            InitializeFormTrace();
            inputSubscription = SetupThrottle();
        }


        private void CheckDatabaseConnection()
        {
            using (SqlConnection connection = new(connectionString))
            {
                try
                {
                    connection.Open();
                    ShowSuccessMessage("Po��czenie z baz� danych zosta�o nawi�zane.");
                }
                catch (SqlException ex)
                {
                    ShowErrorMessage("B��d po��czenia z baz� danych: " + ex.Message);
                }
            }
        }

        private void CheckExcelFileAvailability()
        {
            // Pobierz pe�n� �cie�k� do pliku wykonywalnego aplikacji
            string executablePath = Assembly.GetExecutingAssembly().Location;
            string executableDirectory = System.IO.Path.GetDirectoryName(executablePath);

            string labelDirectory = System.IO.Path.Combine(executableDirectory, "Label"); // Folder "Label" w tym samym katalogu, co plik wykonywalny
            string excelFileName = "label.xlsx";
            string excelFilePath = System.IO.Path.Combine(labelDirectory, excelFileName);

            if (File.Exists(excelFilePath))
            {
                ShowSuccessMessage("Plik Excel jest dost�pny.");
            }
            else
            {
                ShowErrorMessage("Plik Excel nie jest dost�pny w lokalizacji: " + excelFilePath);
            }
        }

        private IDisposable SetupThrottle()
        {
            var syncContext = SynchronizationContext.Current;
            return userInputSubject
                .Throttle(TimeSpan.FromSeconds(ThrottleTimeSeconds))
                .ObserveOn(syncContext)
                .Subscribe(_ => HandleUserInput());
        }

        public void InitializeFormTrace()
        {
            if (textBoxes == null)
            {
                // Zg�o� b��d lub inicjalizuj textBoxes
                return;
            }

            foreach (var textBox in textBoxes)
            {
                textBox.TextChanged += TextBox_TextChanged;
                textBox.KeyUp += TextBox_KeyUp; // Dodaj obs�ug� KeyUp
            }
        }

        private void TextBox_KeyUp(object sender, KeyEventArgs e)
        {
            // Sygnalizowanie wprowadzenia danych przy ka�dym naci�ni�ciu klawisza
            userInputSubject.OnNext(Unit.Default);
        }

        private void HandleUserInput()
        {
            if (Monitor.TryEnter(lockObject))
            {
                try
                {
                    ProcessCurrentTextBox();
                }
                finally
                {
                    Monitor.Exit(lockObject);
                }
            }
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
                MessageBox.Show(message, "B��d", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ProcessCurrentTextBox()
        {
            var currentTextBox = textBoxes[currentTextBoxIndex];
            var text = currentTextBox.Text;

            // Specjalna obs�uga dla textBoxtrace
            if (currentTextBox == textBoxtrace)
            {
                if (text.Length == MaxCharacterCount)
                {
                    pn = text.Substring(13);
                    textBoxPN.Text = pn;
                    textBoxPN.Enabled = false;
                }
                else
                {
                    // Je�li textBoxtrace nie ma 25 znak�w, upewnij, �e textBoxPN jest aktywny
                    textBoxPN.Enabled = true;
                }
            }

            // Sprawd�, czy mo�na przej�� do textBoxPN
            if (currentTextBox == textBoxPN && !textBoxPN.Enabled)
            {
                // Je�li textBoxPN jest wy��lczzoney, przeskocz do nast�pnego pola
                MoveToNextTextBox();
            }
            else
            {
                // Normalne przesuni�cie do nast�pnego pola
                MoveToNextTextBox();
            }
        }
                
        private void MoveToNextTextBox()
        {
            // Sprawdzanie, czy aktualne pole jest puste
            var currentTextBox = textBoxes[currentTextBoxIndex];
            if (string.IsNullOrWhiteSpace(currentTextBox.Text))
            {
                // Je�li pole jest puste, nie przechod� do nast�pnego
                return;
            }

            do
            {
                // Przej�cie do nast�pnego pola tekstowego
                currentTextBoxIndex = (currentTextBoxIndex + 1) % textBoxes.Length;
            }
            while (currentTextBoxIndex != 0 && !textBoxes[currentTextBoxIndex].Enabled); // Pomijaj wy��czone pola

            if (currentTextBoxIndex < textBoxes.Length)
            {
                var nextTextBox = textBoxes[currentTextBoxIndex];
                nextTextBox.Focus();
            }
            else
            {
                MessageBox.Show("Uzupe�ni�e� wszystkie wymagane pola.");
                currentTextBoxIndex = 0; // Resetuj indeks do pocz�tkowego stanu
            }
        }

        private void TextBox_TextChanged(object sender, EventArgs e)
        {
            // Sygnalizowanie wprowadzenia danych r�wnie� w przypadku zmiany tekstu
            userInputSubject.OnNext(Unit.Default);
        }

        private string GenerateLotCode(DateTime date)
        {
            // Logika generowania kodu lotu na podstawie daty
            char yearCode = (char)('A' + (date.Year - 2023) % 26);
            char monthCode = (char)('A' + date.Month - 1);

            // Je�li przekroczono 25 liter alfabetu, zaczynamy u�ywa� cyfr (1 dla literki A, 2 dla B, itd.)
            char dayCode = date.Day <= 25 ? (char)('A' + date.Day - 1) : (char)('1' + date.Day - 26);

            return $"{yearCode}{monthCode}{dayCode}";
        }        

        private async void OpenAndPrintExcelFileHandler(object sender, EventArgs e)
        {
            string trace = textBoxtrace.Text;
            string trace2 = textBoxtrace2.Text;
            string rackQty = textBoxrackqty.Text;
            string rack = textBoxrack.Text;
            string rack2 = textBoxrack2.Text;

            try
            {
                if (trace.Length == MaxCharacterCount)
                {
                    
                    string p_trace = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + rackQty + rack + pn;

                    var (partName, revValue, barcodeValue) = await GetPartNameRevAndBarcodeAsync(pn);
                    rev = revValue;
                    string barcode = barcodeValue;                   
                    GenerateAndSaveQRCode(trace, trace2, pn, rev, rackQty, barcode);
                    GenerateBarcodeAndSave(barcodeValue);                    
                    InsertRecord(pn, DateTime.Now.Date, DateTime.Now.TimeOfDay, rackQty, rack, rack2, trace, trace2, p_trace, barcode);
                    //OpenAndPrintExcelFile(pn, DateTime.Now.Date, DateTime.Now.TimeOfDay, rackQty, rack, rack2, trace, trace2, p_trace, rev, barcode);

                    ShowSuccessMessage("Plik Excel zosta� otwarty i wydrukowany.");
                }
                else if (trace.Length != MaxCharacterCount)
                {
                   string pn = textBoxPN.Text;
                    string p_trace = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + rackQty + rack + pn;

                    var (partName, revValue, barcodeValue) = await GetPartNameRevAndBarcodeAsync(pn);
                    rev = revValue;
                    string barcode = barcodeValue;

                   GenerateAndSaveQRCode(trace, trace2, pn, rev, rackQty, barcode);
                   GenerateBarcodeAndSave(barcodeValue);
                   InsertRecord(pn, DateTime.Now.Date, DateTime.Now.TimeOfDay, rackQty, rack, rack2, trace, trace2, p_trace, barcode);
                   //OpenAndPrintExcelFile(pn, DateTime.Now.Date, DateTime.Now.TimeOfDay, rackQty, rack, rack2, trace, trace2, p_trace, rev, barcode);

                    ShowSuccessMessage("Plik Excel zosta� otwarty i wydrukowany.");
                }
                else
                {
                    ShowErrorMessage("B��d: Nieprawid�owa d�ugo�� ci�gu lub brak danych.");
                }
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Wyst�pi� b��d: " + ex.Message);
            }
        }


        private void GenerateBarcodeAndSave(string barcodeValue)
        {
            try
            {
                MessageBox.Show("Rozpocz�cie generowania kodu kreskowego");

                // U�ywamy BarcodeWriterPixelData zamiast BarcodeWriter
                var barcodeWriter = new BarcodeWriterPixelData
                {
                    Format = BarcodeFormat.CODE_128,
                    Options = new EncodingOptions
                    {
                        Width = 300,
                        Height = 100,
                        Margin = 10
                    }
                };

                // Generowanie danych kodu kreskowego w formie pikseli
                PixelData pixelData = barcodeWriter.Write(barcodeValue);

                // Tworzenie bitmapy z danych pikselowych
                using (Bitmap bitmap = new Bitmap(pixelData.Width, pixelData.Height, PixelFormat.Format32bppRgb))
                {
                    // Lock the bits of the bitmap.
                    BitmapData bitmapData = bitmap.LockBits(new System.Drawing.Rectangle(0, 0, pixelData.Width, pixelData.Height), ImageLockMode.WriteOnly, PixelFormat.Format32bppRgb);
                    try
                    {
                        // Update the bitmap with the data from the PixelData
                        System.Runtime.InteropServices.Marshal.Copy(pixelData.Pixels, 0, bitmapData.Scan0, pixelData.Pixels.Length);
                    }
                    finally
                    {
                        bitmap.UnlockBits(bitmapData);
                    }

                    // Pobierz �cie�k� do pliku wykonywalnego aplikacji
                    string executablePath = Assembly.GetExecutingAssembly().Location;
                    string executableDirectory = System.IO.Path.GetDirectoryName(executablePath);

                    // Okre�l �cie�k� do folderu "Barcode"
                    string barcodeDirectory = System.IO.Path.Combine(executableDirectory, "Barcode");

                    // Utw�rz folder, je�li nie istnieje
                    if (!Directory.Exists(barcodeDirectory))
                    {
                        Directory.CreateDirectory(barcodeDirectory);
                    }

                    // Zapisz plik kodu kreskowego
                    string filePath = System.IO.Path.Combine(barcodeDirectory, $"{barcodeValue}.png");
                    bitmap.Save(filePath, ImageFormat.Png);
                    MessageBox.Show($"Kod kreskowy zosta� zapisany w: {filePath}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"B��d podczas generowania kodu kreskowego: {ex.Message}");
            }
        }

        private void GenerateAndSaveQRCode(string trace, string trace2, string pn, string rev, string rackQty, string barcode)
        {
            try
            {
                MessageBox.Show("Rozpocz�cie generowania kodu QR");

                var barcodeWriter = new BarcodeWriterPixelData
                {
                    Format = BarcodeFormat.QR_CODE,
                    Options = new EncodingOptions
                    {
                        Width = 300,
                        Height = 300,
                        Margin = 10
                    }
                };

                string qrText = string.Empty;
                if (!string.IsNullOrEmpty(trace) && !string.IsNullOrEmpty(trace2))
                {
                    qrText = $"[)>06:AS\"barcode\":PN\"{pn}\":QT\"{rackQty}.000\":RV\"{rev}\":DM\"{DateTime.Now.ToString("ddMMyy")}\":SPHS:PO:LT\"{GenerateLotCode(DateTime.Now)}\":WT\"{trace}\" / \"{trace2}\":PT\"{DateTime.Now.ToString("dd.MM.yy")} {DateTime.Now.TimeOfDay}\"/#{rack} / #{rack2}/{pn}:*[]\"";
                }
                else if (!string.IsNullOrEmpty(trace))
                {
                    qrText = $"[)>06:AS\"barcode\":PN\"{pn}\":QT\"{rackQty}.000\":RV\"{rev}\":DM\"{DateTime.Now.ToString("ddMMyy")}\":SPHS:PO:LT\"{GenerateLotCode(DateTime.Now)}\":WT\"{trace}\":PT\"{DateTime.Now.ToString("dd.MM.yy")} {DateTime.Now.TimeOfDay}\"/#{rack} / #{rack2}/{pn}:*[]\"";
                }

                // Generowanie danych kodu QR w formie pikseli
                PixelData pixelData = barcodeWriter.Write(qrText);

                // Tworzenie bitmapy z danych pikselowych
                using (Bitmap bitmap = new Bitmap(pixelData.Width, pixelData.Height, PixelFormat.Format32bppRgb))
                {
                    // Lock the bits of the bitmap.
                    BitmapData bitmapData = bitmap.LockBits(new System.Drawing.Rectangle(0, 0, pixelData.Width, pixelData.Height), ImageLockMode.WriteOnly, PixelFormat.Format32bppRgb);
                    try
                    {
                        // Update the bitmap with the data from the PixelData
                        System.Runtime.InteropServices.Marshal.Copy(pixelData.Pixels, 0, bitmapData.Scan0, pixelData.Pixels.Length);
                    }
                    finally
                    {
                        bitmap.UnlockBits(bitmapData);
                    }

                    // Zapisz kod QR w odpowiednim folderze
                    string filePath = System.IO.Path.Combine("QRCode", $"{barcode}.png");
                    bitmap.Save(filePath, ImageFormat.Png);
                    MessageBox.Show($"Kod QR zosta� zapisany w: {filePath}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"B��d podczas generowania kodu QR: {ex.Message}");
            }
        }


        private async void InsertRecord(string pn, DateTime date, TimeSpan hour, string rackQty, string rack, string rack2, string trace, string trace2, string pTrace, string barcode)
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

                            // Zamknij DataReader, poniewa� ju� uzyskali�my potrzebne dane
                            reader.Close();

                            // Teraz, kiedy masz PartName i Rev, mo�esz je wstawi� do tabeli Archive
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
                                    ShowSuccessMessage($"Rekord zosta� zarchiwizowany. id_trace = {idTrace}");
                                }
                                catch (Exception ex)
                                {
                                    ShowErrorMessage("B��d podczas archiwizacji: " + ex.Message);
                                }
                            }
                        }
                        else
                        {
                            // Obs�u� przypadek, gdy nie znaleziono informacji dla danego PN w tabeli Database
                            ShowErrorMessage($"Brak informacji w tabeli Database dla PN: {pn}");
                        }
                    }
                }
            }
        }

        private async Task<(string partName, string rev, string barcode)> GetPartNameRevAndBarcodeAsync(string pn)
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

                // Pobierz ostatni� warto�� kolumny "Barcode" z tabeli "Archive" na podstawie PN
                string selectArchiveBarcodeQuery = "SELECT TOP 1 [Barcode] FROM [Archive] ORDER BY Date DESC, Hour DESC;\r\n";

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

            // Generuj now� warto�� kolumny "Barcode" na podstawie poprzedniej
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
            
            // Pobierz pe�n� �cie�k� do pliku wykonywalnego aplikacji
            string executablePath = Assembly.GetExecutingAssembly().Location;
            string executableDirectory = System.IO.Path.GetDirectoryName(executablePath);
            string labelDirectory = System.IO.Path.Combine(executableDirectory, "Label"); // Folder "Label" w tym samym katalogu, co plik wykonywalny
            string excelFileName = "label.xlsx";
            string excelFilePath = System.IO.Path.Combine(labelDirectory, excelFileName);


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
                    worksheet.Range["I2"].Value = "VW"; // Dla 25 znak�w w textBoxtrace
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
                Debug.WriteLine("B��d: " + ex.Message);
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
