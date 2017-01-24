using System;
using System.IO;
using System.Collections.Generic;

namespace RapGEN
{

    class LogLoader : ProgressReporter, IDataLoader<DataRow>
    {
        // Pola:
        private string path;
        private bool doRecursive;

        // Konstruktory:
        public LogLoader(string path, bool doRecursive) : base()
        {
            this.path = path;
            this.doRecursive = doRecursive;
        }
        public LogLoader(string path) : this(path, false) { }

        // Metody:
        private string[] CreateFilesList(string path)
        {
            // Future TODO: Rozwinąć tak by przeszukiwało foldery rekurencyjnie (doRecursive)
            string[] FilesList = Directory.GetFiles(path);
            return FilesList;
        }
        public Queue<DataRow> LoadData()
        {
            string row;
            string[] Files = CreateFilesList(path);
            if (Files.Length == 0) throw new LogLoaderException("Podany folder wejściowy jest pusty!");

            Queue<DataRow> Data = new Queue<DataRow>();
            StreamReader sr = null;
            ProgressMax = Files.Length;

            try
            {
                foreach (string f in Files)
                {
                    if (stopRequired) return Data;

                    sr = new StreamReader(f);
                    row = sr.ReadLine();
                    while (row != null)
                    {
                        Data.Enqueue(DataRow.SplitRowToData(row));
                        row = sr.ReadLine();
                    }
                    ++Progress;
                }
            }
            catch (OutOfMemoryException) { throw new LogLoaderException("Za dużo grzybków w barszcz. ;(\nSpróbuj zrestartować program i wybrać mniejszą ilość danych do rozdzielenia."); }
            catch (DataRow.DataRowException ex) { throw new LogLoaderException(ex.Message); }
            catch (Exception) { throw new LogLoaderException("Coś poszło nie tak. Najprawdopodobniej jeden z zadanych plików ma niepoprawny format."); }
            finally { sr.Close(); }

            return Data;
        }

        // Właściwości i wyjątki:
        [Serializable]
        public sealed class LogLoaderException : ApplicationException
        {
            public LogLoaderException(string msg) : base(msg) { }
        }
    }
}
