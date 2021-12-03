using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Serilog;
using Serilog.Sinks.File;


namespace LibraryOfClasses
{
    /// <summary>
    /// Ищет в *.xls файлах слово
    /// </summary>
    public class SearchInDirectory
    {
        public static List<string> Begin( string wordFound )
        {
            Log.Logger = new LoggerConfiguration( )
                 .WriteTo.File( "logExcel.txt", rollingInterval: RollingInterval.Infinite )
                 .CreateLogger( );
            List<string> resultSearch = new List<string>( ); //Список найденных строк

            string[] files = Directory.GetFiles( Directory.GetCurrentDirectory( ), "*.xls", SearchOption.TopDirectoryOnly );
            ExtractExcelFiletoDateSet Excel = null;// Объект для отображения файла Excel на DataSet

            ConcurrentBag<string> cbResultSearch = new ConcurrentBag<string>( );
            List<Task> bagAddTasks = new List<Task>( );


            foreach (string cfile in files)
            {
                bagAddTasks.Add( Task.Run( ( ) =>
                {
                    //try
                    //{
                        Excel = ExcelToTable( cfile );
                        //ResultToCommon( cbResultSearch, Excel.SearchWordInDataSet( wordFound ) );
                        Excel.SearchWordInDataSet( wordFound );
                    //}
                    //catch (Exception ex)
                    //{
                    //    Log.Fatal( $"Ошибка в потоке - {ex.Message}" );
                    //    Console.WriteLine( ex.Message );
                    //}

                } ));

                //Log.Information( $"Файл-< {Path.GetFileNameWithoutExtension( cfile )} >" );
                //Excel = new ExtractExcelFiletoDateSet(
                //    ConfigurationManager.ConnectionStrings[ "Excel16" ].ConnectionString, cfile );
                //Log.Information( $"Импорт данных из файла-< {Path.GetFileNameWithoutExtension( cfile )} >" );
                //resultSearch.AddRange( Excel.SearchWordInDataSet( wordFound ) );
                //Log.Information( $"Поиск завершен-< {Path.GetFileNameWithoutExtension( cfile )} >" );
            }
            Task.WaitAll( bagAddTasks.ToArray( ) );
           // resultSearch.Add( "Заглушка" );

            return resultSearch;
        }

        private static void ResultToCommon( ConcurrentBag<string> cbResultSearch, IEnumerable<string> list )
        {
            foreach (var item in list)
            {
                cbResultSearch.Add( item );
            }
        }

        private static ExtractExcelFiletoDateSet ExcelToTable( string cfile )
        {
            ExtractExcelFiletoDateSet Excel;
            Log.Information( $"Файл-< {Path.GetFileNameWithoutExtension( cfile )} >" );
            Excel = new ExtractExcelFiletoDateSet(
                ConfigurationManager.ConnectionStrings[ "Excel16" ].ConnectionString, cfile );
            Log.Information( $"Импорт данных из файла-< {Path.GetFileNameWithoutExtension( cfile )} >" );
            //resultSearch.AddRange( Excel.SearchWordInDataSet( wordFound ) );
            Log.Information( $"Поиск завершен-< {Path.GetFileNameWithoutExtension( cfile )} >" );
            return Excel;
        }
    }
}
