using System;
using System.Collections.Generic;
using System.Text;
using Dapper;
using MySql.Data.MySqlClient;
using System.Threading.Tasks;
using System.Linq;
using System.Data;
using FluentFTP;

namespace DataLibrary
{
    public class DataAccess
    {

        static string DBstring = "";
        static string FTP_conn = "";
        static string _userName = "";
        static string _password = "";



        public static void DBmain(int _db)
        {
            if (_db == 1)
            {
                DBstring = "Server=78.57.2.98;Port=3306;Database=k1020packinglist;Uid=teltonikagamyba;Pwd=pN48PYrqf;SslMode = none";// Namai Raspberry Pi
                FTP_conn = "78.57.2.98";// Namai Rasberry Pi
                _userName = "pi";
                _password = "rutuliukas";
            }
            else if (_db == 2)
            {
                DBstring = "Server=ems-projekt.ad.teltonika.lt;Port=3306;Database=k1020packinglist;Uid=teltonikagamyba;Pwd=pN48PYrqf; SslMode = none";// Teltonika PC
                FTP_conn = "ems-projekt.ad.teltonika.lt";// Teltonika Other
                _userName = "Teltonika";
                _password = "teltonikaems";

            }
            
        }

        public static bool ConnToMysql()
        {
            bool result;

            MySqlConnection connection = new MySqlConnection(DBstring);

            try
            {
                connection.Open();
                result = true;
                connection.Close();
            }
            catch
            {
                result = false;
            }
            return result;
        }

        public static async Task<List<T>> LoadData<T, U>(string sql, U parameters)
        {
            using (IDbConnection connection = new MySqlConnection(DBstring))
            {
                var rows = await connection.QueryAsync<T>(sql, parameters);

                return rows.ToList();
            }
        }

        public static Task SaveData<T>(string sql, T parameters)
        {
            using (IDbConnection connection = new MySqlConnection(DBstring))
            {
                return connection.ExecuteAsync(sql, parameters);
            }
        }

        public static FtpClient CreateFtpClient()
        {
            return new FtpClient(FTP_conn, new System.Net.NetworkCredential { UserName = _userName, Password = _password });
        }
    }
}

