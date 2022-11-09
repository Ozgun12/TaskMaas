using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Text;
using Dapper;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;

namespace MaaslariAktar
{
    internal class MaasIslemleri
    {

        private Workbook xlWorkBook;

        public void Run()
        {
            var maaslist = GetMaaslist();

            var excellByte = CreateExcell(maaslist);

            DeleteGecmisMaaasKayitlari(maaslist);
        }
        private void DeleteGecmisMaaasKayitlari(List<Maas> maaslist)
        {
     
            var sql = "sp_EskiMaasSil";
            using var connection = new SqlConnection("SERVER=192.168.3.223;DATABASE=db_a880f6_alokazaik;User Id=sa;Password=Soft2022!!");

            connection.Open();
            var controlDate = DateTime.UtcNow;

            var query = connection.ExecuteAsync(sql,commandType:System.Data.CommandType.StoredProcedure).Result;

        }

        private byte[] CreateExcell(List<Maas> maaslist)
        {
            string dosya_yolu = @"C:\Users\Osman\Desktop\Maaas1.txt";
            FileStream fs = new FileStream(dosya_yolu, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            

            return null;
        }

        private List<Maas> GetMaaslist()
        {
            var sql = "sp_EskiMaasListele";
            using var connection = new SqlConnection("SERVER=192.168.3.223;DATABASE=db_a880f6_alokazaik;User Id=sa;Password=Soft2022!!");

            connection.Open();
          var query =  connection.QueryAsync<Maas>(sql, new
            {
            });

            return query.Result.AsList();
        }
    }

    public class Maas
    {
        public int Id { get; set; }
        public int PersonelId { get; set; }
        public string AdSoyad { get; set; }
        public string TC { get; set; }
        public float YapilanOdeme { get; set; }
        public string MaasOdemeTarihi { get; set; }
        public float KesilenAvans { get; set; }
        public float KesilenHaciz { get; set; }
        public float EkKesinti { get; set; }
    }

    

}
