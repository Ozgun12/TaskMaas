using System;
using System.Data;
using System.Data.SqlClient;

namespace MaaslariAktar
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var massıslemleri = new MaasIslemleri();
            massıslemleri.Run();
            Console.WriteLine("Büyük Özgün");
            Console.ReadLine();
        }
    }
}
