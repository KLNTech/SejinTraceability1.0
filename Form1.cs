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
using static QRCoder.PayloadGenerator.ShadowSocksConfig;
using System.Xml.Linq;
using System.Data;
using System.Linq;

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
        private const int MaxIdleTimeSeconds = 1;
        private bool projectSelectionPending = false;
        private bool isAutoMoveInProgress = false;
        private readonly object lockObject = new();
        private bool isFormOpened = false;

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
      .ObserveOn(SynchronizationContext.Current)  // Zmiana ta powinna rozwi�za� problem
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
                MessageBox.Show(message, "B��d", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string skipAutoMove = null;

        private void HandleTextChanged(string text)
        {
            // ta metoda monitoruje zmiany w tek�cie i podejmuje r�ne dzia�ania
            // w zale�no�ci od d�ugo�ci tekstu i pola tekstowego, kt�re zosta�o zmienione.
            // Je�li tekst ma odpowiedni� d�ugo��, mo�e to oznacza� zako�czenie wprowadzania danych,
            // co inicjuje przeniesienie do nast�pnego pola tekstowego.
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

                    // Przesu� do nast�pnego pola tylko, je�li u�ytkownik sko�czy� wprowadza� ci�g znak�w
                    ThrottleMoveToNextTextBox();
                }
            }
            else if (text.Length != MaxCharacterCount && textBoxes[currentTextBoxIndex] == textBoxtrace)
            {
                projectSelectionPending = true;
                // Przesu� do nast�pnego pola tylko, je�li u�ytkownik sko�czy� wprowadza� ci�g znak�w
                ThrottleMoveToNextTextBox();
            }
        }

        private async Task MoveToNextTextBox()
        {
            //ta metoda zarz�dza procesem automatycznego przechodzenia do nast�pnego pola tekstowego,
            //uwzgl�dniaj�c r�ne warunki i scenariusze, takie jak wyb�r projektu, pomini�cie ruchu automatycznego,
            //sprawdzenie pustego pola tekstowego, obs�uga ostatniego pola, oraz op�nienie przed przej�ciem do nast�pnego pola.
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
            // Metoda ThrottleMoveToNextTextBox odpowiada za kontrolowanie automatycznego
            // przechodzenia do nast�pnego pola tekstowego, ale z ograniczeniem czasowym (throttle)

            Console.WriteLine("ThrottleMoveToNextTextBox called");

            lock (lockObject)
            {
                // Ustawianie blokady (lockObject) w celu zabezpieczenia przed
                // r�wnoczesnym dost�pem wielu w�tk�w do kodu chronionego t� blokad�.
                if (isAutoMoveInProgress)
                {
                    // Sprawdzanie, czy automatyczne przechodzenie do nast�pnego pola (isAutoMoveInProgress)
                    // jest ju� w trakcie. Je�li tak, to metoda ko�czy si�, poniewa� nie mo�na r�wnocze�nie
                    // wykonywa� wielu operacji tego typu.
                    return;
                }

                isAutoMoveInProgress = true;
                // Je�li automatyczne przechodzenie nie jest w trakcie, ustawia flag� na true,
                // aby zablokowa� kolejne wywo�ania tej metody.
            }

            Console.WriteLine($"Before MoveToNextTextBox: TextBoxIndex: {currentTextBoxIndex}");

            // Oczekiwanie na asynchroniczne wykonanie metody MoveToNextTextBox
            await MoveToNextTextBox();

            Console.WriteLine($"After MoveToNextTextBox: TextBoxIndex: {currentTextBoxIndex}");

            lock (lockObject)
            {
                // Zdejmowanie blokady, ustawiaj�c flag� isAutoMoveInProgress na false, co oznacza,
                // �e teraz mo�na ponownie wywo�a� t� metod�.
                isAutoMoveInProgress = false;
            }
        }


        private void ShowMessageBox()
        {
            if (textBoxrackqty.InvokeRequired)
            {
                textBoxrackqty.Invoke(new MethodInvoker(delegate { ShowMessageBox(); }));
            }
            else
            {
                MessageBox.Show("Wprowad� dane przed przej�ciem dalej.");
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
            //SetActiveControl to metoda, kt�ra ustawia fokus(aktywny element, kt�ry reaguje na klawisze klawiatury) na jednym z p�l tekstowych(TextBox)
            //w zale�no�ci od bie��cego indeksu currentTextBoxIndex. Je�eli wywo�anie tej metody zachodzi w w�tku interfejsu u�ytkownika(UI), to fokus ustawiany
            //jest bezpo�rednio.W przeciwnym razie(gdy wywo�anie pochodzi z innego w�tku ni� UI), metoda Invoke jest u�ywana do prze��czenia wykonania na w�tek UI,
            //gdzie nast�pnie ustawiany jest fokus.
            if (textBoxes[currentTextBoxIndex].InvokeRequired)
            {
                textBoxes[currentTextBoxIndex].Invoke((MethodInvoker)delegate
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
            isFormOpened = false; // Okno zosta�o ju� otwarte
        }

        private void ShowProjectSelectionDialog()
        {
            if (!projectSelectionPending || isFormOpened)
            {
                // Je�eli nie oczekuje si� na wyb�r projektu lub okno jest ju� otwarte, nie otwieraj okna.
                return;
            }

            isFormOpened = true; // Ustaw flag�, �eby zapobiec otwarciu okna wi�cej ni� raz
            ProjectSelectionForm projectSelectionForm = new ProjectSelectionForm(); // Create an instance
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

                (string revValue, string barcodeValue) = GetRevAndBarcode(pn);
                string rev = revValue;
                string barcode = barcodeValue;

                try
                {
                    InsertRecord(pn, DateTime.Now.Date, DateTime.Now.TimeOfDay, rackQty, rack, rack2, trace, trace2, p_trace, barcode);
                    OpenAndPrintExcelFile(pn, DateTime.Now.Date, DateTime.Now.TimeOfDay, rackQty, rack, rack2, trace, trace2, p_trace, rev, barcode);

                    ShowSuccessMessage("Plik Excel zosta� otwarty i wydrukowany.");
                }
                catch (Exception ex)
                {
                    ShowErrorMessage("B��d: " + ex.Message);
                }
                finally
                {
                    // Dzia�ania do wykonania po zako�czeniu pr�by otwarcia i wydrukowania pliku Excel.
                }
            }
            else if (trace.Length > 25 || !string.IsNullOrEmpty(rack))
            {
                // Obs�uga przypadku dla instrukcji "traceability" dla VW
                // Tutaj mo�esz doda� kod obs�uguj�cy instrukcj� "traceability"
                // oraz archiwizacj� wynik�w tej instrukcji
            }
            else
            {
                ShowErrorMessage("B��d: Nieprawid�owa d�ugo�� ci�gu lub brak danych.");
            }
        }


        private void InsertRecord(string pn, DateTime date, TimeSpan hour, string rackQty, string rack, string rack2, string trace, string trace2, string pTrace, string barcode)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string insertQuery = "INSERT INTO Archive (PN, Date, Hour, RackQty, Rack, Rack2, Trace, Trace2, PTrace, Barcode) " +
                                     "VALUES (@pn, @date, @hour, @rack_qty, @rack, @rack2, @trace, @trace2, @p_trace, @barcode); " +
                                     "SELECT CAST(SCOPE_IDENTITY() AS INT)";

                using (SqlCommand cmd = new SqlCommand(insertQuery, connection))
                {
                    cmd.Parameters.AddWithValue("@pn", pn);
                    cmd.Parameters.AddWithValue("@date", date);
                    cmd.Parameters.AddWithValue("@hour", hour);
                    cmd.Parameters.AddWithValue("@rack_qty", rackQty);
                    cmd.Parameters.AddWithValue("@rack", rack);
                    cmd.Parameters.AddWithValue("@rack2", rack2);
                    cmd.Parameters.AddWithValue("@trace", trace);
                    cmd.Parameters.AddWithValue("@trace2", trace2);
                    cmd.Parameters.AddWithValue("@p_trace", pTrace);
                    cmd.Parameters.AddWithValue("@barcode", barcode);

                    try
                    {
                        connection.Open();
                        // Wykorzystaj ExecuteScalar, aby uzyska� warto�� Identity dla nowo dodanego rekordu
                        int idTrace = (int)cmd.ExecuteScalar();
                        ShowSuccessMessage($"Rekord zosta� zarchiwizowany. id_trace = {idTrace}");
                    }
                    catch (Exception ex)
                    {
                        ShowErrorMessage("B��d podczas archiwizacji: " + ex.Message);
                    }
                }
            }
        }



        private (string revValue, string barcodeValue) GetRevAndBarcode(string pn)
        {
            string rev = string.Empty;
            string barcode = string.Empty;

            using (SqlConnection connection = new(connectionString))
            {
                connection.Open();
                string query = "SELECT TOP 1 barcode FROM Archive ORDER BY date DESC, hour DESC";
                using SqlCommand cmd = new(query, connection);
                using SqlDataReader reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    barcode = reader["barcode"].ToString();
                }
            }

            string firstPart = barcode[..7];
            string secondPart = barcode[7..];

            if (int.TryParse(secondPart, out int secondPartNumber))
            {
                secondPartNumber++;
                string incrementedSecondPart = secondPartNumber.ToString("D6");
                barcode = firstPart + incrementedSecondPart;
            }

            return (rev, barcode);
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
                    Console.WriteLine("Nie znaleziono arkusza o nazwie 'label'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("B��d: " + ex.Message);
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
