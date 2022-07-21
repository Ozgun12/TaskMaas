using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Text;
using Dapper;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

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
            Microsoft.Office.Interop.Excel.Application xlApp = new
            Microsoft.Office.Interop.Excel.Application();
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);
            foreach (var item in maaslist)
            {
                for (int i = 0; i <= maaslist.Count; i++)
                {
                    xlWorkSheet.Cells[i, 1] = item.TC;
                    xlWorkSheet.Cells[i, 2] = item.AdSoyad;
                    xlWorkSheet.Cells[i, 3] = item.YapilanOdeme;
                    xlWorkSheet.Cells[i, 4] = item.KesilenHaciz;
                    xlWorkSheet.Cells[i, 5] = item.KesilenAvans;
                    xlWorkSheet.Cells[i, 6] = item.EkKesinti;
                }
                
            }
            xlWorkBook.SaveAs("c:\\maas.xlsx", XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

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
