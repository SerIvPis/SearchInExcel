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
            bool isWork = true;
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
                    List<string> resultSearch = null;
                    resultSearch = SearchInDirectory.Begin( wordForSearch );
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
            }
        }

        private static void PrintList( List<string> resultSearch )
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
