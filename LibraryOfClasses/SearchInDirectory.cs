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
using System.Data;
using System.Diagnostics;

namespace LibraryOfClasses
{
    /// <summary>
    /// Ищет в *.xls файлах слово
    /// </summary>
    public class SearchInDirectory
    {
        public static List<string> resultSearchList = new List<string>();

        //public static ConcurrentBag<string> cbResultSearch { get; set; }

        public void Begin( string wordFound )
        {
            Log.Logger = new LoggerConfiguration( )
                 .WriteTo.File( "logExcel.txt", rollingInterval: RollingInterval.Minute )
                 .CreateLogger( );
            string[] files = Directory.GetFiles( Directory.GetCurrentDirectory( ), "*.xls", SearchOption.TopDirectoryOnly );

            List<Task> ltask = new List<Task>( );
            resultSearchList.Clear();
            Parallel.ForEach( files, cfile =>
            {
                //Console.WriteLine($"поток = {Thread.CurrentThread.Name}");
                 try
                 {
                     Log.Information( $"Файл-< {Path.GetFileNameWithoutExtension( cfile )} >" );
                     Stopwatch timer = new Stopwatch( );
                     timer.Start( );
                     ExtractExcelFiletoDateSet Excel = new ExtractExcelFiletoDateSet(
                        ConfigurationManager.ConnectionStrings[ "ExcelODBC" ].ConnectionString, cfile,wordFound );
                     Log.Information( $"Импорт данных из файла-< {Path.GetFileNameWithoutExtension( cfile )} >" );
                     resultSearchList.AddRange( Excel.SearchWordInDataSet( wordFound ) );
                     timer.Stop( );
                     Log.Information( $"Поиск завершен-< {Path.GetFileNameWithoutExtension( cfile )} >" +
                        $"Время поиска: { timer.Elapsed.TotalSeconds:g}сек." +
                        $"" );

                 }
                 catch (Exception e)
                 {
                     Log.Fatal( $"Ошибка в {e.StackTrace} - {e.Message}" );
                     Console.WriteLine( e.Message );
                     //Console.WriteLine( e.StackTrace );

                 }
            } );
                   
               
            
        }
       
                //public static List<string> SearchWordInDataSet( string wordFound )
                //{
                //    //List<string> resultList = new List<string>( );
                //          cbResultSearch = new ConcurrentBag<string>( );
                 
                //           foreach (DataTable curDt in ExcelFileDataSet.Tables)
                //           {
                //                   foreach (DataRow dataRow in curDt.Rows)
                //                   {
                //                       foreach (DataColumn item in curDt.Columns)
                //                       {
                //                           if (dataRow[ item ].ToString( ).Trim( ).ToLowerInvariant( ).Equals( wordFound.Trim( ).ToLowerInvariant( )
                //                               , StringComparison.InvariantCultureIgnoreCase ))
                //                           {
                //                               //cbResultSearch.Add( PrintDataRow( dataRow, curDt, ExcelFileDataSet ) );
                //                               Console.WriteLine($"{PrintDataRow( dataRow, curDt, ExcelFileDataSet )}");
                //                           }
                //                       }
                //                   }
                //           }


                //    return cbResultSearch.ToList<string>();

                //}

                //private static string PrintDataRow( DataRow dataRow, DataTable curDt, DataSet excelFileDataSet )
                //{
                //    string resultStr = null;

                //    if (curDt.TableName.Contains( "список" ))// Для первого листа "Список"
                //    {
                //        resultStr = $" [ {curDt.TableName} ] --->\t" +
                //        $"{dataRow[ "Обозначение" ]}\t" +
                //        $"{dataRow[ "Наименование" ]}";
                //    }
                //    else if (curDt.Columns.Contains( "Инв# № подл#" ))// Для групповой спецификации
                //    {
                //        resultStr = $" [ {curDt.TableName} ] --->\t" +
                //       $"{dataRow[ "F7" ]}\t" +
                //       $"{dataRow[ "Инв# № дубл#" ]}\t";
                //    }
                //    else  // Для одиночной спецификации
                //    {
                //        resultStr = $"[ {curDt.TableName} ] --->\t" +
                //       $"{dataRow[ "Обозначение" ]}\t" +
                //       $"{dataRow[ "Наименование" ]}\t" +
                //       $"{dataRow[ "Кол#" ]}";
                //    }
                //    return resultStr;
                //}
                
    }
}
