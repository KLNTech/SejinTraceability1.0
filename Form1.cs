using System;
using System.Data.SqlClient;
using System.Reactive.Linq;
using System.Reactive.Subjects;
using System.Reactive.Concurrency;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using System.IO;
using TextBox = System.Windows.Forms.TextBox;
using Action = System.Action;

namespace SejinTraceability
{
    public partial class straceabilitysystem : Form
    {
        private readonly string connectionString;
        private TextBox[] textBoxes;
        private int currentTextBoxIndex = 0;
        private Subject<string> inputSubject = new Subject<string>();
        private IDisposable inputSubscription;
        private readonly SynchronizationContext _syncContext;

        public straceabilitysystem()
        {
            InitializeComponent();
            _syncContext = WindowsFormsSynchronizationContext.Current;
            string connectionString = SejinTraceability.Properties.Settings.Default.ConnectionString;

            // SprawdŸ, czy ConnectionString jest prawid³owy
            if (!string.IsNullOrEmpty(connectionString))
            {
                // Ustaw ConnectionString w obiekcie SqlConnection
                SqlConnection connection = new SqlConnection(connectionString);

                // Teraz mo¿esz u¿yæ obiektu 'connection' do wykonywania operacji na bazie danych.
            }
            else
            {
                // Obs³u¿ przypadki, gdy ConnectionString jest pusty lub null
                Console.WriteLine("B³¹d: ConnectionString nie zosta³ prawid³owo skonfigurowany.");
            }
            if (string.IsNullOrEmpty(connectionString))
            {
                // Jeœli connectionString jest puste lub null, zwróæ b³¹d
                throw new InvalidOperationException("ConnectionString property has not been initialized.");
            }

            textBoxes = new TextBox[] { textBoxtrace, textBoxtrace2, textBoxPN, textBoxrackqty, textBoxrack, textBoxrack2 };

            inputSubscription = Observable.Merge(
                Observable.FromEventPattern(textBoxtrace, "TextChanged").Select(pattern => ((TextBox)pattern.Sender).Text),
                Observable.FromEventPattern(textBoxtrace2, "TextChanged").Select(pattern => ((TextBox)pattern.Sender).Text),
                Observable.FromEventPattern(textBoxrackqty, "TextChanged").Select(pattern => ((TextBox)pattern.Sender).Text),
                Observable.FromEventPattern(textBoxrack, "TextChanged").Select(pattern => ((TextBox)pattern.Sender).Text),
                Observable.FromEventPattern(textBoxrack2, "TextChanged").Select(pattern => ((TextBox)pattern.Sender).Text)
            )
            .Throttle(TimeSpan.FromSeconds(1))
            .DistinctUntilChanged()
            .ObserveOn(SynchronizationContext.Current) // U¿yj obecnego kontekstu synchronizacji
            .Subscribe(text => HandleTextChanged(text));
            textBoxtrace2.Leave += TextBoxTrace2_Leave;
            textBoxtrace2.Validated += TextBoxTrace2_Validated;
        }

        private bool isAutoUpdate = false; // Flaga wskazuj¹ca, czy aktualizacja pola jest automatyczna

        private void UpdateTextBoxPN(string text)
        {
            isAutoUpdate = true; // Ustawiamy flagê na true przed automatyczn¹ aktualizacj¹
            textBoxPN.Text = text;
            isAutoUpdate = false; // Przywracamy flagê do wartoœci false po zakoñczeniu automatycznej aktualizacji
        }
                private void HandleTextChanged(string text)
        {
            if (text.Length >= 25 && !isAutoUpdate)
            {
                string pn = text.Substring(13);
                UpdateTextBoxPN(pn);
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
                MessageBox.Show("Wprowadzono dane do ostatniego pola (textBoxrack).");
            }

            ActiveControl = textBoxes[currentTextBoxIndex];
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            inputSubscription?.Dispose();
            base.OnHandleDestroyed(e);
        }


       private void TextBox_TextChanged(object sender, EventArgs e)
{
    System.Windows.Forms.TextBox textBox = (System.Windows.Forms.TextBox)sender;

    if ((textBox == textBoxtrace || textBox == textBoxtrace2) && textBox.Text.Length >= 25)
    {
        string pn = textBox.Text.Substring(13);
        textBoxPN.Text = string.IsNullOrEmpty(textBoxPN.Text) ? pn : textBoxPN.Text;
        if (textBox == textBoxtrace2 && textBoxtrace2.Text.Length >= 25)
        {
            // SprawdŸ, czy dane zosta³y wprowadzone do textBoxTrace2 przed przeskoczeniem
            MoveToNextTextBox();
        }
        else if (textBox == textBoxtrace && textBox.Text.Length == 25)
        {
            // Jeœli d³ugoœæ tekstu w textBoxTrace wynosi dok³adnie 25 znaków, pomijaj przeskakiwanie do textBoxPN
        }
        else
        {
            MoveToNextTextBox();
        }
    }
    else if (textBox == textBoxrackqty && !string.IsNullOrEmpty(textBoxrackqty.Text))
    {
        MoveToNextTextBox();
    }
    else if (textBox == textBoxrack && !string.IsNullOrEmpty(textBoxrack.Text))
    {
        MoveToNextTextBox();
    }
    else if (textBox == textBoxrack2 && textBoxrack.Text.Length >= 25 && textBoxtrace2.Text.Length >= 25)
    {
        MoveToNextTextBox();
    }
}
        private void TextBoxTrace2_Leave(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            if (textBox.Text.Length >= 25)
            {
                textBoxrackqty.Focus();
            }
        }
        private void TextBoxTrace2_Validated(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            if (textBox.Text.Length >= 25)
            {
                MoveToNextTextBox();
            }
        }

        private void TextBoxTrace2_TextChanged(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;

            if (textBox == textBoxtrace2 && textBox.Text.Length >= 25)
            {
                textBoxrackqty.Focus(); // Przeskakujemy do textBoxrackqty po wprowadzeniu 25 znaków w textBoxtrace2
            }
        }

        private void InsertRecord(string pn, DateTime date, TimeSpan hour, string rackQty, string rack, string trace, string pTrace, string revValue, string barcodeValue)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string insertQuery = "INSERT INTO Archive (pn, date, hour, rack_qty, rack, trace, p_trace) " +
                                     "VALUES (@pn, @date, @hour, @rack_qty, @rack, @trace, @ptrace)";

                using (SqlCommand cmd = new SqlCommand(insertQuery, connection))
                {
                    cmd.Parameters.AddWithValue("@pn", pn);
                    cmd.Parameters.AddWithValue("@date", date);
                    cmd.Parameters.AddWithValue("@hour", hour);
                    cmd.Parameters.AddWithValue("@rack_qty", rackQty);
                    cmd.Parameters.AddWithValue("@rack", rack);
                    cmd.Parameters.AddWithValue("@trace", trace);
                    cmd.Parameters.AddWithValue("@p_trace", pTrace);

                    try
                    {
                        connection.Open();
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Rekord zosta³ zarchiwizowany.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("B³¹d podczas archiwizacji: " + ex.Message);
                    }
                }
            }
        }

        private (string rev, string updatedBarcode) GetRevAndBarcode(string pn)
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

        private void OpenAndPrintExcelFile(string pn, DateTime date, TimeSpan hour, string rackQty, string rack, string trace, string p_trace, string revValue, string barcodeValue)
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
                    worksheet.Range["G10"].Value = revValue;
                    worksheet.Range["I6"].Value = barcodeValue;
                    worksheet.Range["A14"].Value = p_trace;
                    worksheet.Range["A18"].Value = rackQty;
                    worksheet.Range["E18"].Value = date;
                    worksheet.Range["C21"].Value = hour;

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

        private void PrintAndArchiveClick(object sender, EventArgs e)
        {
            string trace = textBoxtrace.Text;
            string rackQty = textBoxrackqty.Text;
            string rack = textBoxrack.Text;
            string trace2 = textBoxtrace2.Text;
            string rack2 = textBoxrack2.Text;



            if (trace.Length == 25)
            {
                //trace 13 znakowy jest dla projektów PM, jest to unikalna d³ugoœæ dla tego projektu
                string pn = trace.Substring(13);
                string p_trace = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + rackQty + rack + pn;

                (string revValue, string barcodeValue) result = GetRevAndBarcode(pn);
                string revValue = result.revValue;
                string barcodeValue = result.barcodeValue;

                InsertRecord(pn, DateTime.Now.Date, DateTime.Now.TimeOfDay, rackQty, rack, trace, p_trace, revValue, barcodeValue);
                OpenAndPrintExcelFile(pn, DateTime.Now.Date, DateTime.Now.TimeOfDay, rackQty, rack, trace, p_trace, revValue, barcodeValue);
            }
            else if (trace.Length > 25 || !string.IsNullOrEmpty(rack))
            {
                // Obs³uga przypadku dla instrukcji "traceability" dla VW
                // Tutaj mogê dodaæ kod obs³uguj¹cy instrukcjê "traceability"
                // oraz archiwizacjê wyników tej instrukcji
            }
            else
            {
                //mo¿liwoœæ rozszerzania kodu o dodatkowe projekty
                MessageBox.Show("B³¹d: Nieprawid³owa d³ugoœæ ci¹gu lub brak danych.");
            }
        }




        //private void checkBox1_CheckedChanged(object sender, EventArgs e)
        //{
        //zmiana bazy dabych do zapisu na tabele stripping
        //w momencie ponownego zeskanowania etykiety rekord jest usuwany
        //to ma byc mini WMS na to co jest aktualnie w reworku
        //}



    }
}