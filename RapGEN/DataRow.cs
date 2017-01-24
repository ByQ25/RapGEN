using System;
using System.Windows.Forms;

namespace RapGEN
{
    public struct DataRow
    {
        // Konstrutory:
        public DataRow(string ip, string login, string date, string task, ushort serverResp, uint bytesNumber)
        {
            this.IP = ip;
            this.Login = login;
            this.Date = date;
            this.Task = task;
            this.ServerResp = serverResp;
            this.DataTransferred = bytesNumber;
        }
        public DataRow(string ip, string date, string task, ushort serverResp, uint bytesNumber) : this(ip, "", date, task, serverResp, bytesNumber) { }

        // Metody:
        public static DataRow SplitRowToData(string row)
        {
            string ip, login, date, task, strToConvert;
            ushort serverResp;
            uint bytesNumber;
            int[] Indexes =
            { // Pobranie indeksów do podziału wiersza na poszczególne dane.
                row.IndexOf("-"), // Indexes[0] <- położenie '-' w ciągu znaków
                row.IndexOf("["), // Indexes[1] <- położenie '[' w ciągu znaków
                row.IndexOf("]"), // Indexes[2] <- położenie ']' w ciągu znaków
                row.IndexOf("\""), // Indexes[3] <- położenie pierwszego '"' w ciągu znaków
                row.Length - 1, // Indexes[4] <- (po przejściu poniższego while-a) położenie ostatniego '"' w ciągu znaków
            };

            // Przykładowa linia logu: 62.210.170.165 -  [03/May/2016:00:00:01 +0200] "GET /pz?MP_module=main&MP_action=siwz_tab&noticeIdentity=1184435898 HTTP/1.1" 200 17066
            try
            {
                while (row[Indexes[4]] != '"') --Indexes[4];
                ip = row.Substring(0, Indexes[0] - 1);
                if (Indexes[1] == Indexes[0] + 3) login = "";
                else login = row.Substring(Indexes[0] + 2, Indexes[1] - Indexes[0] - 3);
                date = row.Substring(Indexes[1] + 1, Indexes[2] - Indexes[1] - 1);
                task = row.Substring(Indexes[3] + 1, Indexes[4] - Indexes[3] - 1);
                strToConvert = row.Substring(Indexes[4] + 2, 3);
                if (strToConvert == "-") serverResp = 0;
                else serverResp = Convert.ToUInt16(strToConvert);
                strToConvert = row.Substring(Indexes[4] + 6);
                if (strToConvert == "-") bytesNumber = 0;
                else bytesNumber = Convert.ToUInt32(strToConvert);
            }
            catch (ArgumentOutOfRangeException exOutOfRange)
            {
                throw new DataRowException(exOutOfRange.Message + "\n\nPrzyczyną problemu jest najprawdopodobniej niewłaściwe formatowanie danych w plikach logu. Skontaktuj się z twórcą programu.");
            }
            catch (FormatException formatEx)
            {
                throw new DataRowException(formatEx.Message);
            }
            return new DataRow(ip, login, date, task, serverResp, bytesNumber);
        }

        // Właściwości:
        public string IP { get; private set; }
        public string Login { get; private set; }
        public string Date { get; private set; }
        public string Task { get; private set; }
        public ushort ServerResp { get; private set; }
        public uint DataTransferred { get; private set; }

        // Wyjątki:
        [Serializable]
        public sealed class DataRowException : ApplicationException
        {
            public DataRowException(string msg) : base(msg) { }
        }
    }
}
