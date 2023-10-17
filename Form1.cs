using System;
using System.Data.SqlClient;
using System.Reactive.Linq;
using System.Reactive.Subjects;
using System.Reactive.Concurrency;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace SejinTraceability
{
    public partial class straceabilitysystem : Form, IObserver<string>
    {
        private readonly string connectionString;
        private System.Windows.Forms.TextBox[] textBoxes;
        private IDisposable inputSubscription;
        private int currentTextBoxIndex = 0;
        private DateTime lastKeyPressTime = DateTime.Now;
        private System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
        private Subject<string> inputSubject = new Subject<string>();

        public straceabilitysystem()
        {
            InitializeComponent();
            connectionString = SejinTraceability.Properties.Settings.Default.ConnectionString;
            textBoxes = new System.Windows.Forms.TextBox[] { textBoxtrace, textBoxtrace2, textBoxPN, textBoxrackqty, textBoxrack, textBoxrack2 };

            foreach (var textBox in textBoxes)
            {
                textBox.TextChanged += TextBoxes_TextChanged;
            }

            this.Load += (sender, e) => {
                SetActiveControl(textBoxtrace);
            };
        }

        private void TextBoxes_TextChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox textBox = (System.Windows.Forms.TextBox)sender;
            string text = textBox.Text;

            if (!string.IsNullOrEmpty(text))
            {
                inputSubject.OnNext(text);
                MoveToNextTextBox();
            }
        }

        // ...
        private void SetActiveControl(System.Windows.Forms.TextBox targetTextBox)
        {
            // Ustaw aktywny kontrolnik w polu textBox na w¹tku g³ównym
            this.Invoke(new System.Action(() =>
            {
                ActiveControl = targetTextBox;
            }));
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

            this.Invoke(new System.Action(() =>
            {
                ActiveControl = textBoxes[currentTextBoxIndex];
            }));
        }

        public void OnCompleted()
        {
          
        }
        public void OnError(Exception error)
        {
     
        }
        public void OnNext(string value)
        {
          
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

       /* private void PrintLabel(string pn, DateTime date, TimeSpan hour, string rackQty, string rack, string trace, string p_trace, string revValue, string barcodeValue)
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
                    string qrString = "[)06:AS" + barcodeValue + ":PN" + pn + ":QT...blablabla";
                    string p = (trace.Length == 13) ? "PM" : "VW";

                    worksheet.Range["A7"].Value = pn;
                    worksheet.Range["I2"].Value = p;
                    worksheet.Range["G10"].Value = revValue;
                    worksheet.Range["I6"].Value = qrString;
                    worksheet.Range["A14"].Value = p_trace;
                    worksheet.Range["A18"].Value = rackQty;
                    worksheet.Range["E18"].Value = date;
                    worksheet.Range["C21"].Value = hour;
                    // ... pozosta³e wartoœci
                    worksheet.PrintOut();
                }
                else
                {
                    Console.WriteLine("Nie znaleziono arkusza o nazwie 'label'.");
                }

                workbook.Close(true, excelFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("B³¹d: " + ex.Message);
            }
            finally
            {
                excelApp.Quit();
            }
        }
*/
        
        private (string rev, string updatedBarcode) GetRevAndBarcode(string pn)
        {
            string rev = string.Empty;
            string barcode = string.Empty;

            /*   using (var context = new DatabaseContext())
               {
                   var databaseRecord = context.Database.FirstOrDefault(item => item.PN == pn);
                   if (databaseRecord != null)
                   {
                       rev = databaseRecord.Rev;
                   }
               }
            */
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

            // Utwórz aplikacjê Excel
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            try
            {
                // Otwórz istniej¹cy plik Excel
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
                    // Wype³nij odpowiednie komórki danymi
                    worksheet.Range["A7"].Value = pn;
                    worksheet.Range["I2"].Value = "PM"; // Dla przyk³adu ustawiamy wartoœæ PM
                    worksheet.Range["G10"].Value = revValue;
                    worksheet.Range["I6"].Value = barcodeValue;
                    worksheet.Range["A14"].Value = p_trace;
                    worksheet.Range["A18"].Value = rackQty;
                    worksheet.Range["E18"].Value = date;
                    worksheet.Range["C21"].Value = hour;
                    // ... pozosta³e wartoœci

                    // Drukuj plik Excel na domyœlnej drukarce
                    worksheet.PrintOut();

                    // Zamknij plik Excel
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
                // Zamknij aplikacjê Excel
                excelApp.Quit();

                // Zwalnianie zasobów COM
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
            


            if (trace.Length == 13)
            {
                //trace 13 znakowy jest dla projektów PM, jest to unikalna d³ugoœæ dla tego projektu
                string pn = trace.Substring(15, 8);
                string p_trace = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + rackQty + rack + pn;

                (string revValue, string barcodeValue) result = GetRevAndBarcode(pn);
                string revValue = result.revValue;
                string barcodeValue = result.barcodeValue;

                InsertRecord(pn, DateTime.Now.Date, DateTime.Now.TimeOfDay, rackQty, rack, trace, p_trace, revValue, barcodeValue);
                OpenAndPrintExcelFile(pn, DateTime.Now.Date, DateTime.Now.TimeOfDay, rackQty, rack, trace, p_trace, revValue, barcodeValue);
            }
            else if (trace.Length > 13 || !string.IsNullOrEmpty(rack))
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

    
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //zmiana bazy dabych do zapisu na tabele stripping
            //w momencie ponownego zeskanowania etykiety rekord jest usuwany
            //to ma byc mini WMS na to co jest aktualnie w reworku
        }

      
    }
}

