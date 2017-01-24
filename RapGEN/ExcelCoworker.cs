using System;
using System.Drawing;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace RapGEN
{
    class ExcelCoworker
    {
        // Pola:
        private string pgLabel;
        private int pgValue, pgMax;

        // Kontruktory:
        public ExcelCoworker(ref string pgLabel, ref int pgValue, ref int pgMax)
        {
            this.pgLabel = pgLabel;
            this.pgValue = pgValue;
            this.pgMax = pgMax;
        }

        // Metody:
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("W trakcie zwalniania zasobów wystąpił wyjątek:  " + ex.Message, "Błąd przy zwalnianiu uchwytu Excela", MessageBoxButtons.OK);
            }
            finally
            {
                GC.Collect();
            }
        }
        public void FormatExcelTable(ref Excel.Worksheet exCellWorksheet, int lastRowNr)
        {
            // Formatowanie pierwszego wiersza:
            Excel.Range formatRange = exCellWorksheet.get_Range("a1", "f1");
            formatRange.Interior.Color = ColorTranslator.ToOle(Color.Orange);
            formatRange.Font.Bold = true;
            formatRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            formatRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            formatRange.WrapText = true;
            formatRange.RowHeight = 32;
            formatRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            // Formatowanie obramowania wierszy:
            pgMax = lastRowNr;
            for (int i = 3; i < lastRowNr + 1; ++i)
            {
                formatRange = exCellWorksheet.get_Range("a" + i, "f" + i);
                formatRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                pgValue += 1;
                pgLabel = "Postęp - etap 2, formatowanie tabeli Excela (" + Convert.ToString(Math.Round((double)(pgValue * 100 / pgMax))) + "%):";
            }

            // Formatowanie kolumn:
            formatRange = exCellWorksheet.get_Range("a1", "a" + lastRowNr);
            formatRange.ColumnWidth = 16;
            formatRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            formatRange = exCellWorksheet.get_Range("b1", "b" + lastRowNr);
            formatRange.ColumnWidth = 32;
            formatRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            formatRange = exCellWorksheet.get_Range("c1", "c" + lastRowNr);
            formatRange.ColumnWidth = 32;
            formatRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            formatRange = exCellWorksheet.get_Range("d1", "d" + lastRowNr);
            formatRange.ColumnWidth = 128;
            formatRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            formatRange = exCellWorksheet.get_Range("e1", "e" + lastRowNr);
            formatRange.ColumnWidth = 16;
            formatRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            formatRange = exCellWorksheet.get_Range("f1", "f" + lastRowNr);
            formatRange.ColumnWidth = 16;
            formatRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            // Formatowanie obramowania zewnętrznego:
            formatRange = exCellWorksheet.get_Range("a1", "f" + lastRowNr);
            formatRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
        }
        [Obsolete("This method of direct data export to Excel is incredibly slow. Use rather ConvertXmlToXlsx method.")]
        public void ExpToExcel(ref Excel.Worksheet exCellWorksheet, Queue<DataRow> Data)
        {
            // Pierwszy wiersz:
            exCellWorksheet.Name = "Dane z logów";
            exCellWorksheet.Cells[1, "A"] = "IP";
            exCellWorksheet.Cells[1, "B"] = "Login";
            exCellWorksheet.Cells[1, "C"] = "Data";
            exCellWorksheet.Cells[1, "D"] = "Polecenie";
            exCellWorksheet.Cells[1, "E"] = "Odpowiedź serwera";
            exCellWorksheet.Cells[1, "F"] = "Rozmiar odpowiedzi (bajty)";

            // Uzupełnianie danych w arkuszu:
            pgMax = Data.Count;
            DataRow dr;
            for (int i = 2; i < pgMax + 2; ++i)
            {
                dr = Data.Dequeue();
                exCellWorksheet.Cells[i, "A"].Value = dr.IP;
                exCellWorksheet.Cells[i, "B"].Value = dr.Login;
                exCellWorksheet.Cells[i, "C"].Value = dr.Date;
                exCellWorksheet.Cells[i, "D"].Value = dr.Task;
                exCellWorksheet.Cells[i, "E"].Value = dr.ServerResp;
                exCellWorksheet.Cells[i, "F"].Value = dr.DataTransferred;
            }
        }
        public static void ConvertXmlToXlsx(string xmlFilePath, string outputFilePath)
        {
            // Inicjalizacja Excela:
            Excel.Application exCellApp = new Excel.Application();
            if (exCellApp == null)
                throw new ExcelCoworkerException("Wystąpił problem z programem Excel. Możliwe, że nie został on poprawnie zainstalowany na tym komputerze.");
            Excel.Workbook exCellWorkbook = exCellApp.Workbooks.Open(xmlFilePath);
            object misValue = System.Reflection.Missing.Value;
            try
            {
                exCellWorkbook.SaveAs(outputFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                exCellWorkbook.Close(true, misValue, misValue);
                exCellApp.Quit();
            }
            catch (Exception) { throw new ExcelCoworkerException("Błąd w trakcie próby zapisania pliku *.xlsx."); }
            finally
            {
                // Zwolnienie zasobów (Excel):
                releaseObject(exCellApp);
                releaseObject(exCellWorkbook);
            }
        }

        // Wyjątki:
        [Serializable]
        public sealed class ExcelCoworkerException : ApplicationException
        {
            public ExcelCoworkerException(string msg) : base(msg) { }
        }

        // Unused snippets:
        /*
         * // Inicjalizacja Excela:
         *              Excel.Application exCellApp = new Excel.Application();
         *              if (exCellApp == null)
         *              {
         *                  MessageBox.Show("Wystąpił problem z programem Excel. Możliwe, że nie został on poprawnie zainstalowany na tym komputerze.");
         *                  err = true;
         *                  return;
         *              }
         *              Excel.Workbook exCellWorkbook = exCellApp.Workbooks.Add();
         *              exCellWorkbook.Worksheets[3].Delete();
         *              exCellWorkbook.Worksheets[2].Delete();
         *              Excel.Worksheet exCellWorksheet = (Excel.Worksheet)exCellWorkbook.Worksheets.get_Item(1);
         *              object misValue = System.Reflection.Missing.Value;
         *
         *               // Operacje na Excelu:
         *              FormatExcelTable(ref exCellWorksheet, Data.Count + 1);
         *              pgValue = 0; // Zerowanie paska postępu.
         *              ExpToExcel(ref exCellWorksheet, Data);
         *
         *              // Zapisanie pliku *.xlsx
         *              try
         *              {
         *                  string path = OutputPathTB.Text;
         *                  if (OutputPathTB.Text[path.Length - 1] != '\\') path += @"\";
         *                  exCellWorkbook.SaveAs(path + "(RapGEN) Logi.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
         *                  exCellWorkbook.Close(true, misValue, misValue);
         *                  exCellApp.Quit();
         *              }
         *              catch (Exception ex) { MessageBox.Show(ex.Message, "Błąd w trakcie próby zapisania pliku", MessageBoxButtons.OK); err = true; }
         *
         *              // Zwolnienie zasobów:
         *              releaseObject(exCellApp);
         *              releaseObject(exCellWorkbook);
         *              releaseObject(exCellWorksheet);
         *          }
         */
    }
}