using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;

namespace RapGEN
{
    class ToXmlExporter : ProgressReporter, IXmlExporter<DataRow>, IDisposable
    {
        // Pola
        private FileStream fs;
        private StreamWriter sw;
        private bool disposedValue = false; // To detect redundant disposing calls

        // Konstruktory:
        public ToXmlExporter(string path) : base()
        {
            try
            {
                this.fs = new FileStream(path, FileMode.Create, FileAccess.Write);
                this.sw = new StreamWriter(fs);
                
            }
            catch (Exception ex)
            {
                this.fs = null;
                this.sw = null;
                MessageBox.Show(ex.Message, "Błąd w trakcie próby otwarcia pliku XML do zapisu", MessageBoxButtons.OK);
            }
        }
        public ToXmlExporter(FileStream fs, StreamWriter sw) : base()
        {
            this.fs = fs;
            this.sw = sw;
            this.Progress = 0;
        }
        public ToXmlExporter(StreamWriter sw) : this(null, sw) { }

        // Metody:
        public void WriteBeginning()
        {
            sw.WriteLine("<?xml version=\"1.0\"?>\n");
            sw.WriteLine("<ex:Workbook xmlns:ex=\"urn:schemas-microsoft-com:office:spreadsheet\">");
            // Definiowanie styli:
            sw.WriteLine("\t<ex:Styles>");
            sw.WriteLine("\t\t<ex:Style ex:ID=\"Default\" ex:Name=\"Normal\">");
            sw.WriteLine("\t\t\t<ex:Alignment ex:Vertical=\"Center\"/>");
            sw.WriteLine("\t\t\t<ex:Font ex:FontName=\"Czcionka tekstu podstawowego\" ex:CharSet=\"238\" ex:Family=\"Swiss\" ex:Size=\"11\" ex:Color=\"#000000\"/>");
            sw.WriteLine("\t\t\t<ex:NumberFormat ex:Format=\"0\"/>");
            sw.WriteLine("\t\t</ex:Style>");
            sw.WriteLine("\t\t<ex:Style ex:ID=\"w1\">");
            sw.WriteLine("\t\t\t<ex:Alignment ex:Horizontal=\"Center\" ex:Vertical=\"Center\" ex:WrapText=\"1\"/>");
            sw.WriteLine("\t\t\t<ex:Font ex:Bold=\"1\"/>");
            sw.WriteLine("\t\t</ex:Style>");
            // Formatowanie obramowania:
            sw.WriteLine("\t\t<ex:Style ex:ID=\"cL1\">");
            sw.WriteLine("\t\t\t<ex:Borders>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Left\" ex:LineStyle=\"Continuous\" ex:Weight=\"3\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Right\" ex:LineStyle=\"Continuous\" ex:Weight=\"2\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Top\" ex:LineStyle=\"Continuous\" ex:Weight=\"3\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Bottom\" ex:LineStyle=\"Continuous\" ex:Weight=\"1\"/>");
            sw.WriteLine("\t\t\t</ex:Borders>");
            sw.WriteLine("\t\t\t<ex:Font ex:Bold=\"1\"/>");
            sw.WriteLine("\t\t\t<ex:Interior ex:Color=\"#FFA500\" ex:Pattern=\"Solid\"/>");
            sw.WriteLine("\t\t\t<ex:Alignment ex:Horizontal=\"Center\" ex:Vertical=\"Center\" ex:WrapText=\"1\"/>");
            sw.WriteLine("\t\t</ex:Style>");
            sw.WriteLine("\t\t<ex:Style ex:ID=\"cM1\">");
            sw.WriteLine("\t\t\t<ex:Borders>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Left\" ex:LineStyle=\"Continuous\" ex:Weight=\"2\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Right\" ex:LineStyle=\"Continuous\" ex:Weight=\"2\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Top\" ex:LineStyle=\"Continuous\" ex:Weight=\"3\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Bottom\" ex:LineStyle=\"Continuous\" ex:Weight=\"1\"/>");
            sw.WriteLine("\t\t\t</ex:Borders>");
            sw.WriteLine("\t\t\t<ex:Font ex:Bold=\"1\"/>");
            sw.WriteLine("\t\t\t<ex:Interior ex:Color=\"#FFA500\" ex:Pattern=\"Solid\"/>");
            sw.WriteLine("\t\t\t<ex:Alignment ex:Horizontal=\"Center\" ex:Vertical=\"Center\" ex:WrapText=\"1\"/>");
            sw.WriteLine("\t\t</ex:Style>");
            sw.WriteLine("\t\t<ex:Style ex:ID=\"cR1\">");
            sw.WriteLine("\t\t\t<ex:Borders>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Left\" ex:LineStyle=\"Continuous\" ex:Weight=\"2\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Right\" ex:LineStyle=\"Continuous\" ex:Weight=\"3\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Top\" ex:LineStyle=\"Continuous\" ex:Weight=\"3\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Bottom\" ex:LineStyle=\"Continuous\" ex:Weight=\"1\"/>");
            sw.WriteLine("\t\t\t</ex:Borders>");
            sw.WriteLine("\t\t\t<ex:Font ex:Bold=\"1\"/>");
            sw.WriteLine("\t\t\t<ex:Interior ex:Color=\"#FFA500\" ex:Pattern=\"Solid\"/>");
            sw.WriteLine("\t\t\t<ex:Alignment ex:Horizontal=\"Center\" ex:Vertical=\"Center\" ex:WrapText=\"1\"/>");
            sw.WriteLine("\t\t</ex:Style>");
            sw.WriteLine("\t\t<ex:Style ex:ID=\"cL2\">");
            sw.WriteLine("\t\t\t<ex:Borders>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Left\" ex:LineStyle=\"Continuous\" ex:Weight=\"3\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Right\" ex:LineStyle=\"Continuous\" ex:Weight=\"2\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Top\" ex:LineStyle=\"Continuous\" ex:Weight=\"1\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Bottom\" ex:LineStyle=\"Continuous\" ex:Weight=\"1\"/>");
            sw.WriteLine("\t\t\t</ex:Borders>");
            sw.WriteLine("\t\t</ex:Style>");
            sw.WriteLine("\t\t<ex:Style ex:ID=\"cM2\">");
            sw.WriteLine("\t\t\t<ex:Borders>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Left\" ex:LineStyle=\"Continuous\" ex:Weight=\"2\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Right\" ex:LineStyle=\"Continuous\" ex:Weight=\"2\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Top\" ex:LineStyle=\"Continuous\" ex:Weight=\"1\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Bottom\" ex:LineStyle=\"Continuous\" ex:Weight=\"1\"/>");
            sw.WriteLine("\t\t\t</ex:Borders>");
            sw.WriteLine("\t\t</ex:Style>");
            sw.WriteLine("\t\t<ex:Style ex:ID=\"cR2\">");
            sw.WriteLine("\t\t\t<ex:Borders>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Left\" ex:LineStyle=\"Continuous\" ex:Weight=\"2\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Right\" ex:LineStyle=\"Continuous\" ex:Weight=\"3\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Top\" ex:LineStyle=\"Continuous\" ex:Weight=\"1\"/>");
            sw.WriteLine("\t\t\t\t<ex:Border ex:Position=\"Bottom\" ex:LineStyle=\"Continuous\" ex:Weight=\"1\"/>");
            sw.WriteLine("\t\t\t</ex:Borders>");
            sw.WriteLine("\t\t</ex:Style>");
            sw.WriteLine("\t</ex:Styles>");
            sw.Flush();
            fs.Flush();
        }
        public void InsertFirstRow(byte worksheetNr)
        {
            sw.WriteLine("\t<ex:Worksheet ex:Name=\"Dane z logów ({0})\">", worksheetNr);
            sw.WriteLine("\t\t<ex:Table>");
            // Ustawienie szerokości kolumn:
            sw.WriteLine("\t\t\t<ex:Column ex:Width=\"100\"/>");
            sw.WriteLine("\t\t\t<ex:Column ex:Width=\"195\"/>");
            sw.WriteLine("\t\t\t<ex:Column ex:Width=\"195\"/>");
            sw.WriteLine("\t\t\t<ex:Column ex:Width=\"770\"/>");
            sw.WriteLine("\t\t\t<ex:Column ex:Width=\"195\"/>");
            sw.WriteLine("\t\t\t<ex:Column ex:Width=\"195\"/>");
            // Wpisanie pierwszego wiersza:
            sw.WriteLine("\t\t\t<ex:Row ex:Height=\"32\" ex:StyleID=\"w1\">");
            sw.WriteLine("\t\t\t\t<ex:Cell ex:StyleID=\"cL1\">");
            sw.WriteLine("\t\t\t\t\t<ex:Data ex:Type=\"String\">IP</ex:Data>");
            sw.WriteLine("\t\t\t\t</ex:Cell>");
            sw.WriteLine("\t\t\t\t<ex:Cell ex:StyleID=\"cM1\">");
            sw.WriteLine("\t\t\t\t\t<ex:Data ex:Type=\"String\">Login</ex:Data>");
            sw.WriteLine("\t\t\t\t</ex:Cell>");
            sw.WriteLine("\t\t\t\t<ex:Cell ex:StyleID=\"cM1\">");
            sw.WriteLine("\t\t\t\t\t<ex:Data ex:Type=\"String\">Data</ex:Data>");
            sw.WriteLine("\t\t\t\t</ex:Cell>");
            sw.WriteLine("\t\t\t\t<ex:Cell ex:StyleID=\"cM1\">");
            sw.WriteLine("\t\t\t\t\t<ex:Data ex:Type=\"String\">Polecenie</ex:Data>");
            sw.WriteLine("\t\t\t\t</ex:Cell>");
            sw.WriteLine("\t\t\t\t<ex:Cell ex:StyleID=\"cM1\">");
            sw.WriteLine("\t\t\t\t\t<ex:Data ex:Type=\"String\">Kod odpowiedzi</ex:Data>");
            sw.WriteLine("\t\t\t\t</ex:Cell>");
            sw.WriteLine("\t\t\t\t<ex:Cell ex:StyleID=\"cR1\">");
            sw.WriteLine("\t\t\t\t\t<ex:Data ex:Type=\"String\">Rozmiar odpowiedzi (bajty)</ex:Data>");
            sw.WriteLine("\t\t\t\t</ex:Cell>");
            sw.WriteLine("\t\t\t</ex:Row>");
            sw.Flush();
            fs.Flush();
        }
        public void InsertRow(params object[] values)
        {
            string cellDataType;
            sw.WriteLine("\t\t\t<ex:Row>");
            for (int i = 0; i < values.Length; ++i)
            {
                if (i == 0) sw.WriteLine("\t\t\t\t<ex:Cell ex:StyleID=\"cL2\">");
                else if (i == values.Length - 1) sw.WriteLine("\t\t\t\t<ex:Cell ex:StyleID=\"cR2\">");
                else sw.WriteLine("\t\t\t\t<ex:Cell ex:StyleID=\"cM2\">");
                cellDataType = values[i].GetType().ToString().Substring(7);
                if (cellDataType == "String")
                    sw.WriteLine("\t\t\t\t\t<ex:Data ex:Type=\"String\">{0}</ex:Data>", values[i].ToString());
                else
                    sw.WriteLine("\t\t\t\t\t<ex:Data ex:Type=\"Number\">{0}</ex:Data>", values[i].ToString());
                sw.WriteLine("\t\t\t\t</ex:Cell>");
                sw.Flush();
                fs.Flush();
            }
            sw.WriteLine("\t\t\t</ex:Row>");
        }
        public void WriteWorksheetEnding()
        {
            sw.WriteLine("\t\t</ex:Table>");
            sw.WriteLine("\t</ex:Worksheet>");
            sw.Flush();
            fs.Flush();
        }
        public void WriteEnding()
        {
            sw.WriteLine("</ex:Workbook>");
            sw.Flush();
            fs.Flush();
        }
        public void Export(Queue<DataRow> Data)
        {
            byte worksheetNr = 1;
            ProgressMax = Data.Count;

            WriteBeginning();
            InsertFirstRow(worksheetNr++);
            foreach (DataRow dr in Data)
            {
                ++Progress; // Obsługa paska postępu:
                if (stopRequired) goto Ending;
                InsertRow(dr.IP, dr.Login, dr.Date, dr.Task, dr.ServerResp, dr.DataTransferred);
               
                // Rozdzielenie danych na arkusze
                if (Progress % 1048575 == 0)
                {
                    WriteWorksheetEnding();
                    InsertFirstRow(worksheetNr++);
                }
            }
            Ending:
            if (Progress % 1048575 != 0) WriteWorksheetEnding();
            WriteEnding();
        }
        public void DeleteCreatedFile()
        {
            sw.Close();
            if (fs != null) File.Delete(fs.Name);
            else throw new ToXmlExporterException("Plik został wcześniej zamknięty lub usunięty.");
        }
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    sw.Dispose();
                    fs.Dispose();  
                }
                sw.Close();
                fs.Close();
                disposedValue = true;
            }
        }
        public void Dispose()
        {
            Dispose(true);
        }

        // Wyjątki:
        [Serializable]
        public sealed class ToXmlExporterException : ApplicationException
        {
            public ToXmlExporterException(string msg) : base(msg) { }
        }
    }
}
