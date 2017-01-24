using System;
using System.Windows.Forms;
using Microsoft.Win32; //Registry
using System.Collections.Generic;
using System.Threading;
using System.Drawing;

// Background pattern was downloaded from www.subtlepatterns.com

namespace RapGEN
{
    public partial class RapGEN_MainWin : Form
    {
        // Pola:
        private byte stage;
        private bool isCanceled;
        private DialogResult dr;
        private IDataLoader<DataRow> logLoad;
        private IXmlExporter<DataRow> xmlExporter;
        private Thread ComputationalThread;

        // Metody:
        public RapGEN_MainWin()
        {
            InitializeComponent();
            dr = DialogResult.None;
            progBarLabel1.Text = "Postęp";
            UpdateProgressBar(0, 100);
            isCanceled = false;
            // ToolTips
            ToolTip OutputPathTT = new ToolTip(), OutputPathLabTT = new ToolTip();
            OutputPathTT.SetToolTip(this.OutputPathTB, "Domyślna lokalizacja: Pulpit.");
            OutputPathLabTT.SetToolTip(this.OutputPathLabel, "Domyślna lokalizacja: Pulpit.");
            // Odfocusowanie.
            this.ActiveControl = null;
            Select(false, false);
            // Ustawienie domyślnej ścieżki wyjściowej na Pulpit.
            OutputPathTB.Text = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders").GetValue("Desktop").ToString();
            saveFileDialog1.InitialDirectory = OutputPathTB.Text;
            OutputPathTB.Text += "\\(RapGEN) Logi.xlsx";
        }
        public void UpdateProgressBar(int value)
        {
            progressBar1.Value = value;
            progressBar1.Refresh();
        }
        public void UpdateProgressBar(int value, string description)
        {
            progBarLabel1.Text = description;
            progressBar1.Value = value;
            progressBar1.Refresh();
        }
        public void UpdateProgressBar(int value, int valueMax)
        {
            progressBar1.Maximum = valueMax;
            progressBar1.Value = value;
            progressBar1.Refresh();
        }
        private void ResetControls()
        {
            GenButton.Text = "GENERUJ!";
            GenButton.BackColor = Color.YellowGreen;
            GenButton.ForeColor = Color.DarkGreen;
            GenButton.Enabled = true;
            InputPathTB.Enabled = true;
            this.ControlBox = true;
            UpdateProgressBar(0, 100);
            mainTimer.Stop();
        }
        public void ToggleIntoSafeMode()
        {
            this.ControlBox = false;
            GenButton.BackColor = DefaultBackColor;
            GenButton.ForeColor = Color.Black;
            GenButton.Enabled = false;
        }
        private void GenButtonAction()
        {

            string xmlFilePath = OutputPathTB.Text;
            if (OutputPathTB.Text.Contains("xlsx")) xmlFilePath = xmlFilePath.Substring(0, xmlFilePath.Length - 4) + "xml";

            // Pierwszy etap (stage) procesu - wczytywanie logów:
            stage = 1;
            logLoad = new LogLoader(InputPathTB.Text);
            Queue<DataRow> Data = null;
            try
            {
                Data = logLoad.LoadData();
                if (isCanceled) goto Ending;

                // Drugi etap (stage) procesu - export danych do pliku *.xml:
                stage = 2;
                xmlExporter = new ToXmlExporter(xmlFilePath);
                xmlExporter.Export(Data);
                if (isCanceled) goto Ending;

                // (Opcjonalnie) Trzeci etap procesu - otwarcie wygenerowanego pliku *.xml w Excelu działającym w tle oraz zapis danych do *.xlsx:
                stage = 3;
                if (OutputPathTB.Text.Contains(".xlsx"))
                {
                    ExcelCoworker.ConvertXmlToXlsx(xmlFilePath, OutputPathTB.Text);
                    xmlExporter.DeleteCreatedFile();
                }
                if (isCanceled) goto Ending;

                // Czwarty etap procesu - resetowanie kontrolek i czyszczenie zasobów:
                stage = 4;

                Ending:
                if (isCanceled && stage > 1) xmlExporter.DeleteCreatedFile();
            }
            catch (ApplicationException exc)
            {
                MessageBox.Show(exc.Message, "Błąd", MessageBoxButtons.OK);
                MessageBox.Show("Próba wygenerowania raportu zakończona niepowodzeniem.", "Rezultat", MessageBoxButtons.OK);
            }
            finally
            {
                // Czyszczenie zasobów:
                Data.Clear();
                Data = null;
                lock (new object())
                    if (xmlExporter is IDisposable) (xmlExporter as IDisposable).Dispose();
                xmlExporter = null;
                GC.Collect();
                isCanceled = false;
            }
        }

        // Obsługa zdarzeń:
        private void BrowseButton1_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.SelectedPath = "";
            dr = folderBrowserDialog1.ShowDialog();
            if (dr == DialogResult.OK)
                InputPathTB.Text = folderBrowserDialog1.SelectedPath;
        }
        private void BrowseButton2_Click(object sender, EventArgs e)
        {
            saveFileDialog1.FileName = "";
            dr = saveFileDialog1.ShowDialog();
            if (dr == DialogResult.OK)
                OutputPathTB.Text = saveFileDialog1.FileName;
        }
        private void GenButton_Click(object sender, EventArgs e)
        {
            if (GenButton.Text == "GENERUJ!")
            {
                progBarLabel1.Text = "Postęp:";
                GenButton.BackColor = Color.DarkOrange;
                GenButton.ForeColor = Color.Black;
                GenButton.Text = "ANULUJ!";
                InputPathTB.Enabled = false;

                // Tworzenie i uruchamianie osobnego wątku:
                ComputationalThread = new Thread(GenButtonAction);
                ComputationalThread.IsBackground = true; // Sprawi, że wątek zakończy się wraz z zamknięciem okna programu.
                ComputationalThread.Start();

                mainTimer.Start();
            }
            else
            {
                if (ProcessCanceled != null) ProcessCanceled(this, EventArgs.Empty); // Wywołanie zdarzenia.
                isCanceled = true;
                ComputationalThread.Join();
                ResetControls();
                progBarLabel1.Text = "Postęp: przerwano generowanie raportu.";
            }
        }
        private void InputPathTB_Changed(object sender, EventArgs e)
        {
            if (InputPathTB.Text.Length != 0)
            {
                GenButton.Enabled = true;
                GenButton.BackColor = Color.YellowGreen;
            }
            else
            {
                GenButton.BackColor = DefaultBackColor;
                GenButton.Enabled = false;
            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            // Obsługa paska postępu:
            byte pProg = 0;
            switch (stage)
            {
                case 1:
                    pProg = logLoad.PercentageProgress;
                    UpdateProgressBar(pProg, string.Format("Postęp: etap 1 - podział danych na kolumny: ({0}%)", pProg));
                    break;
                case 2:
                    pProg = xmlExporter.PercentageProgress;
                    UpdateProgressBar(pProg, string.Format("Postęp: etap 2 - eksport danych do formatu *.xml: ({0}%)", pProg));
                    break;
                case 3:
                    ToggleIntoSafeMode();
                    progBarLabel1.Text = "Postęp: trwa konwersja do formatu *.XLSX, poczekaj...";
                    break; 
                case 4:
                    ResetControls();
                    progBarLabel1.Text = "Postęp: zakończono generowanie raportu.";
                    MessageBox.Show("Plik raportu został wygenerowany.", "Rezultat", MessageBoxButtons.OK);
                    break;
                default:
                    UpdateProgressBar(0, 100);
                    break;
            }
        }

        // Zdarzenia:
        public static event EventHandler ProcessCanceled;

        // Wyjątki:
        [Serializable]
        public sealed class RapGENException : ApplicationException
        {
            public RapGENException(string msg) : base(msg) { }
        }
    }
}