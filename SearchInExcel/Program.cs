using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
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
using System.Data.Odbc;

namespace SearchInExcel
{
    

    class Program
    {
        
        public static void Main( string[] args )
        {
            Log.Logger = new LoggerConfiguration( )
                .WriteTo.File( "logExcel.txt", rollingInterval: RollingInterval.Minute )
                .CreateLogger( );
            Console.WriteLine( new string( '-', 100 ) );
            Console.WriteLine($" +++++++ Start ++++++ ");
            Console.WriteLine( new string( '-', 100  )); 
            Console.WriteLine( "Введите децимальный номер для поиска:" );
            Console.WriteLine();
            string wordForSearch = Console.ReadLine( );
            Console.WriteLine( new string( '-', 100 ));

            // Создаем потокозащещенную коллекцию для хранения импортированных таблиц из excel
            //ConcurrentBag<DataTable> listTables = new ConcurrentBag<DataTable>();
            List<DataTable> listTables = new List<DataTable>();

            //Просмотр файлов в каталоге
            string[] files = Directory.GetFiles( Path.Combine(Directory.GetCurrentDirectory( ), "Files"), "*.xls", SearchOption.TopDirectoryOnly );
            string connectionString =
               ConfigurationManager.ConnectionStrings[ "ExcelODBC" ].ConnectionString;
            ParallelOptions parallelOptions = new ParallelOptions( );
            Parallel.ForEach( files, cfile =>
            {
                try
                {
                    // Создаем объект для импорта из Excel файла
                    ExcelDAL excelDAL = new ExcelDAL( );
                    Log.Information( $"Файл-< {Path.GetFileNameWithoutExtension( cfile )} >" );
                   // Console.WriteLine( $"Файл-< {Path.GetFileNameWithoutExtension( cfile )} >" );
                    Stopwatch timer = new Stopwatch( );
                    timer.Start( );
                    OdbcConnectionStringBuilder build = new OdbcConnectionStringBuilder( connectionString )
                    {
                        [ "Dbq" ] = cfile
                    };
                    //Подключение к файлу
                    excelDAL.OpenConnection( build.ConnectionString );
                    //Импорт в DataTable
                    foreach (var item in excelDAL.GetAllSheets( cfile ))
                    {
                        listTables.Add( item );
                    }
                    //Отключение от файла
                    excelDAL.CloseConnection( );

                    Log.Information( $"Импорт данных из файла-< {Path.GetFileNameWithoutExtension( cfile )} >" );
                    timer.Stop( );
                    Log.Information( $"Импорт < {Path.GetFileNameWithoutExtension( cfile )} >\tзавершен " +
                       $"\t[ { timer.Elapsed.TotalSeconds:g}сек. ]"  );
                    Console.WriteLine( $"Импорт < {Path.GetFileNameWithoutExtension( cfile ),15} > завершен " +
                       $"[ { timer.Elapsed.TotalSeconds:g}сек. ]" );
                }
                catch (Exception e)
                {
                    Log.Fatal( $"Ошибка в {e.StackTrace} - {e.Message}" );
                    Console.WriteLine( e.Message );
                }
            } );
            Console.WriteLine( new string( '-', 100 ) );
            Console.WriteLine( "Начало поиска" );
            Console.WriteLine( new string( '-', 100 ) );

            Stopwatch timer_1 = new Stopwatch( );
            timer_1.Start( );
            //Поиск слова в списке всех импортированных таблицах
            PrintList( SearchWordInDataSet( listTables, wordForSearch ) );
            Console.WriteLine( $"Окончание поиска: { timer_1.Elapsed.TotalSeconds:g}сек." );
            timer_1.Stop( );
            Log.Information( $"Поиск завершен-<>" +
                      $"Время поиска: { timer_1.Elapsed.TotalSeconds:g}сек." +
                      $"" );
            //}
            Console.ReadLine( );
        }

        /*
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
                     //PrintList( resultSearch );
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
                 //Console.WriteLine( e.Message );
                 //Console.ReadLine( );
             }
        }
        */

        public static List<string> SearchWordInDataSet( IEnumerable<DataTable> listTables, string wordFound )
        {
            List<string> resultList = new List<string>( );
            try
            {
                foreach (DataTable curDt in listTables)
                {
                    foreach (DataRow dataRow in curDt.Rows)
                    {
                        foreach (DataColumn item in curDt.Columns)
                        {
                            if (dataRow[ item ].ToString( ).Trim( ).ToLowerInvariant( ).Equals( wordFound.Trim( ).ToLowerInvariant( )
                                , StringComparison.InvariantCultureIgnoreCase ))
                            {
                                resultList.Add( PrintDataRow( dataRow, curDt ) );
                                //Console.WriteLine($"{PrintDataRow( dataRow, curDt, ExcelFileDataSet )}");
                            }
                        }
                    }
                }
                return resultList;
            }
            catch (Exception e)
            {

                Log.Fatal( $"Ошибка в {e.StackTrace} - {e.Message}" );
                Console.WriteLine( $"{e.Message} метод {e.TargetSite}" );
                return resultList;

            }
        }


        private static string PrintDataRow( DataRow dataRow, DataTable curDt )
        {
            string resultStr = null;

            if (curDt.TableName.Contains( "список" ))// Для первого листа "Список"
            {
                resultStr = $"[ {curDt.TableName} ] --->\t" +
                $"{dataRow[ "Обозначение" ]}\t" +
                $"{dataRow[ "Наименование" ]}";
            }
            else if (curDt.Columns.Contains( "Инв# № подл#" ))// Для групповой спецификации
            {
                resultStr = $"[ {curDt.TableName} ] --->\t" +
               $"{dataRow[ "F7" ]}\t" +
               $"{dataRow[ "Инв# № дубл#" ]}\t";
            }
            else  // Для одиночной спецификации
            {
                resultStr = $"[ {curDt.TableName} ] --->\t" +
               $"{dataRow[ "Обозначение" ]}\t" +
               $"{dataRow[ "Наименование" ]}\t" +
               $"{dataRow[ "Кол#" ]}";
            }

            return resultStr;
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
