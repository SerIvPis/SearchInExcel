using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LibraryOfClasses;
using System.Configuration;
using System.IO;
using System.Data;
using System.Diagnostics;
using Serilog.Sinks.File;
using Serilog;

namespace SearchInExcel
{
    

    class Program
    {
        static void Main( string[] args )
        {
            Log.Information( "Старт приложения" );
            string wordForSearch = null;
            string stopWord = "й";
            //string cfile = @"d:\6ДМ-086.xls";
            //ExtractExcelFiletoDateSet Excel = new ExtractExcelFiletoDateSet(
            //  ConfigurationManager.ConnectionStrings[ "ExcelODBC" ].ConnectionString, cfile );
            bool isWork = true;
            SearchInDirectory search = new SearchInDirectory( );
            try
            {
                 do
                 {
                     Console.WriteLine( new string( '-', 100 ) );
                     Console.WriteLine( "Введите децимальный номер для поиска:" );
                     wordForSearch = Console.ReadLine( );
                     Stopwatch timer = new Stopwatch( );
                     timer.Start( );
                     Log.Information( $"Ищем слово -[{wordForSearch}]" );
                     IEnumerable<string> resultSearch = SearchInDirectory.resultSearchList;
                    search.Begin( wordForSearch );
                     
                     //resultSearch = SearchInDirectory.Begin( wordForSearch );
                     PrintList( resultSearch );
                     timer.Stop( );
                     Console.WriteLine($"Время поиска: {timer.Elapsed.TotalSeconds:g} секунд");
                     timer.Reset( );
                     Console.WriteLine( $"Нажмите \"{stopWord}\" чтобы завершить работу или любую клавишу чтобы продолжить..." );
                     if (Console.ReadLine( ).ToLowerInvariant( ) == stopWord)
                     {
                         isWork = false;
                     }
                 } while (isWork);
             }
             catch (Exception e)
             {
                 Log.Fatal($"Ошибка в {e.StackTrace} - {e.Message}" );
                 Console.WriteLine( e.Message );
                Console.ReadLine( );
             }
        }

        private static void PrintList( IEnumerable<string> resultSearch )
        {
            Console.WriteLine( new string( '*', 100 ) );
            foreach (var item in resultSearch)
            {
                Console.WriteLine(item);
            }
            Console.WriteLine( new string( '*', 100 ) );
        }
    }
}
